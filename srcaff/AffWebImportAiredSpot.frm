VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmWebImportAiredSpot 
   Caption         =   "Web Import Aired Spots"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AffWebImportAiredSpot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmcTerminate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7275
      Top             =   4920
   End
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7755
      Top             =   5085
   End
   Begin VB.Timer DoEventsTimer 
      Left            =   720
      Top             =   4920
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   5040
      Width           =   2685
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   5040
      Width           =   2685
   End
   Begin VB.ListBox lbcMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      ItemData        =   "AffWebImportAiredSpot.frx":08CA
      Left            =   120
      List            =   "AffWebImportAiredSpot.frx":08CC
      TabIndex        =   0
      Top             =   450
      Width           =   8175
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   4920
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5550
      FormDesignWidth =   8460
   End
   Begin VB.Label lbcWebType 
      Caption         =   "Production Website"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lacTitle2 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5790
   End
End
Attribute VB_Name = "frmWebImportAiredSpot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private smDate As String     'Import Date
Private imNumberDays As Integer
Private imVefCode As Integer
Private imAdfCode As Integer
Private imAllClick As Integer
Private imImporting As Integer
Private imTerminate As Integer
'Private hmMsg As Integer
Private hmTo As Integer
Private hmFrom As Integer
Private cprst As ADODB.Recordset
Private tmCPDat() As DAT
Private smWebImports As String
Private smWebWorkStatus As String
Private lmTotalHeaders As Long
Private lmTotalSpots As Long
Private lmTotalActivityLogs As Long
Private smEarliestSpottDate As String
Private imEOF As Integer
Private smPledgeByEvent As String

Private tmAstInfo() As ASTINFO
Private tmAirSpotInfo() As AIRSPOTINFO
Private tmTempAirSpotInfo() As AIRSPOTINFO
Private tmweblCode As String
Private tmweblType As String
Private tmweblattCode As String
Private tmweblIPAddr As String
Private tmweblPCName As String
Private tmweblCallLetters As String
Private tmweblVehicleName As String
Private tmwebUserName As String
Private tmwebPostDay As String
Private tmweblDate As String
Private tmWeblTime As String
Private smWebSpots As String
Private smWebHeaders As String
Private smWebLogs As String
Private smFileName As String
Private smStatus As String
Private smMsg1 As String
Private smMsg2 As String
Private smDTStamp As String
Private imNextPassNeeded As Integer
Private imImportByAstCodeAgain As Integer
Private lmResolvedBytime As Long
Private tmCsiFtpInfo As CSIFTPINFO
Private imNoSpotsAired As Integer
Private hmAst As Integer
Private smRecordsArray() As String
Const cmOneSecond As Long = 1000
Private imOldImpLayout As Integer
Private lat_recs As ADODB.Recordset
Private rst_Cpf As ADODB.Recordset
Private lst_rst As ADODB.Recordset
Private lmOrigCpfCode As Long
Const cmPathForgLogMsg As String = "WebImportLog.Txt"
Private Const NODATE As String = "1970-01-01"
Private Const MESSAGEBLACK As Long = 0
Private Const MESSAGERED As Long = 255
Private Const MESSAGEGREEN As Long = 39680
'Dan 8/27/18 removed because auto import on cloud failed
'Private myErrors As CLogger
Private myEnt As CENThelper
Private bmMGBypass As Boolean
'8862 Dan I need to read the source in the import file.  If it's an 'auto update' vendor, create the agreement if it doesn't exist.
Private tmAutoUpdateVendors() As VendorInfo

Private Sub cmdCancel_Click()
    If imImporting Then
        imTerminate = True
    End If
    If Not igAutoImport Then
        Unload frmWebImportAiredSpot
    End If

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
    Dim slPathFileName As String
    Dim slResult As String
    Dim ilRet As Integer
    Dim llRet As Long
    Dim llCnt As Long
    Dim slTemp As String
    Dim slResponse As String
    Dim blResponse As Boolean
    
    On Error GoTo ErrHand
    If imImporting Then
        Exit Sub
    End If
    
    DoEvents
    cmdImport.Enabled = False
    DoEvents
    
    lgSTime5 = timeGetTime
    
    lgCount1 = 0
    lgCount2 = 0
    lgCount3 = 0
    lgCount4 = 0
    lgCount5 = 0
    lgCount6 = 0
    lgCount7 = 0
    lgCount8 = 0
    lgCount9 = 0
    lgCount10 = 0
    lgCount11 = 0

    lgTtlTime1 = 0
    lgTtlTime2 = 0
    lgTtlTime3 = 0
    lgTtlTime4 = 0
    lgTtlTime5 = 0
    lgTtlTime6 = 0
    lgTtlTime7 = 0
    lgTtlTime8 = 0
    lgTtlTime9 = 0
    lgTtlTime10 = 0
    lgTtlTime11 = 0
    lgTtlTime12 = 0
    lgTtlTime13 = 0
    lgTtlTime14 = 0
    lgTtlTime15 = 0
    lgTtlTime16 = 0
    lgTtlTime17 = 0
    lgTtlTime18 = 0
    lgTtlTime19 = 0
    lgTtlTime20 = 0
    lgTtlTime21 = 0
    lgTtlTime22 = 0
    lgTtlTime23 = 0
    lgTtlTime24 = 0
    
    
    'MsgBox "Init FTP being called"
    ilRet = mInitFTP()
     
    'Debug
    '    MsgBox "Init FTP returned " & ilRet
    '    MsgBox "mFtpWebFileToServer being called"
    '    ilRet = mFtpWebFileToServer("WebHeaders.txt")
    '    MsgBox "mFtpWebFileToServer returned " & ilRet
    '    Exit Function
    'Debug End
    
    SetResults "Gathering Info...", 0
    gLogMsg CStr(lmTotalActivityLogs) & " activity logs imported.", "WebImportLog.Txt", False
    ilRet = gPopAttInfo()
    If Not ilRet Then
        ilRet = ilRet
    End If
    mBuildFileNames
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    Screen.MousePointer = vbHourglass
    lmTotalHeaders = 0
    lmTotalSpots = 0
    lmTotalActivityLogs = 0
    imImporting = True
    
    
    On Error GoTo 0
    'D.S. 09/06/17
    If Not gTestAccessToWebServer() Then
        gMsgBox "WARNING!" & vbCrLf & vbCrLf & _
               "Web Server Access Error: The Affiliate System does not have access to the web server or the web server is not responding." & vbCrLf & vbCrLf & _
        "No data will be exported to the web site." & vbCrLf & _
        "No data will be imported from the web site." & vbCrLf & _
        "Sign off system immediately and contact system administrator.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'D.S. 02/19/20 Moved the below 3 lines down from Form Load to insure a web connection is made before launcing the task monitor
    tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
    tmcSetTime.Enabled = True
    gUpdateTaskMonitor 1, "ASI"

    If sgWebSiteNeedsUpdating = "True" Then
        imImporting = False
        gLogMsg "Web Version: " & sgWebSiteVersion & " does not agree with Affiliate Web Version: " & sgWebSiteExpectedByAffiliate & " No Imports Are Allowed", "WebImportLog.Txt", False
        gMsgBox "Web Version: " & sgWebSiteVersion & " does not agree with Affiliate Web Version: " & sgWebSiteExpectedByAffiliate & sgCRLF & sgCRLF & "          No Imports Are Allowed Until Corrected." & sgCRLF & sgCRLF & "                       Call Counterpoint!"
        Screen.MousePointer = vbDefault
        Unload frmWebImportAiredSpot
        Exit Sub
    End If


    If (StrComp(sgCommand, "/m", vbTextCompare) <> 0) Then
        Call mWaitForWebLock
        If imTerminate Then
            Screen.MousePointer = vbDefault
            cmdCancel.Enabled = True
            imImporting = False
            SetResults "Import was canceled.", 0
            gLogMsg "The import Process was terminiated by user.", "WebImportLog.Txt", False
            Exit Sub
        End If
    
        gLogMsg "", "WebImportLog.Txt", False
        gLogMsg "** Starting Web Import Process **", "WebImportLog.Txt", False
        ' First retrieve the two files from the web server we need to import.
        gLogMsg "Requesting web site to export spots.", "WebImportLog.Txt", False
        SetResults "Requesting web site to export spots.", 0
        ' JD - The time it takes the web site to export has increased more than 5 minutes.
        '      We had to change to the new dispatch method instead of waiting for the results.
        If Not gExecExtStoredProc(smWebSpots, "ExportSpots.exe", False, False) Then
            Close #hmTo
            SetResults "FAIL: Unable to instruct Web Server to export spots.", RGB(255, 0, 0)
            gLogMsg "ERROR: IMPORT - Unable to instruct Web Server to export spots..", "WebImportLog.Txt", False
            Screen.MousePointer = vbDefault
            cmdCancel.Enabled = True
            imImporting = False
            Call gEndWebSession("WebImportLog.Txt")
            Exit Sub
        End If
        ' JD - Here's the call that waits for the results from the change above.
        ilRet = mExCheckWebWorkStatus(smWebSpots)
        If ilRet = True Then
            gLogMsg "Web import file created successfully.", "WebImportLog.Txt", False
            SetResults "Web import file created successfully.", 0
        Else
            If StrComp(Trim$(smMsg1), "WARN: An export is already running. Commit needed. Operation aborted.") = 0 Then
                gLogMsg "Warning: An import is currently running. Please check back later.", "WebImportLog.Txt", False
                SetResults "Warning: An import is currently running. Please check back later.", RGB(255, 0, 0)
                imTerminate = True
                Screen.MousePointer = vbDefault
                cmdCancel.Enabled = True
                imImporting = False
                Call gEndWebSession("WebImportLog.Txt")
                Exit Sub
            End If
            
            slTemp = Left$(Trim$(smMsg1), 35)
            If StrComp(Trim$(slTemp), "ERROR: Export Count does not match.") = 0 Then
                gLogMsg "   " & smMsg1, "WebImportLog.Txt", False
                'SetResults "Warning: An export is currently running. Please check back later.", RGB(255, 0, 0)
                imTerminate = True
                Screen.MousePointer = vbDefault
                cmdCancel.Enabled = True
                imImporting = False
                Call gEndWebSession("WebImportLog.Txt")
                Exit Sub
            Else
                gLogMsg "Error: Web create import file Failed.", "WebImportLog.Txt", False
                SetResults "Web create import file Failed.", RGB(255, 0, 0)
                imTerminate = True
                Screen.MousePointer = vbDefault
                cmdCancel.Enabled = True
                imImporting = False
                Call gEndWebSession("WebImportLog.Txt")
                Exit Sub
            End If
        End If
    
        SetResults "Requesting web site to export headers.", 0
        gLogMsg "Requesting web site to export headers.", "WebImportLog.Txt", False
                
        If Not gSendCmdToWebServer("ExportHeaders.dll", smWebHeaders) Then
            Close #hmTo
            SetResults "FAIL: Unable to instruct Web Server to export headers.", RGB(255, 0, 0)
            gLogMsg "ERROR: IMPORT - Unable to instruct Web Server to export headers.", "WebImportLog.Txt", False
            Screen.MousePointer = vbDefault
            cmdCancel.Enabled = True
            imImporting = False
            Call gEndWebSession("WebImportLog.Txt")
            Exit Sub
        End If
        SetResults "Requesting web site to export activity logs.", 0
        gLogMsg "Requesting web site to export activity logs.", "WebImportLog.Txt", False
        If Not gSendCmdToWebServer("ExportWebL.dll", smWebLogs) Then
            Close #hmTo
            SetResults "FAIL: Unable to instruct Web Server to export activity logs.", RGB(255, 0, 0)
            gLogMsg "ERROR: IMPORT - Unable to instruct Web Server to export activity logs.", "WebImportLog.Txt", False
            Screen.MousePointer = vbDefault
            cmdCancel.Enabled = True
            imImporting = False
            Call gEndWebSession("WebImportLog.Txt")
            Exit Sub
        End If
    
        SetResults "Receiving headers from web site.", 0
        gLogMsg "Receiving headers from web site.", "WebImportLog.Txt", False
        smWebImports = gSetPathEndSlash(smWebImports, True)
        slPathFileName = smWebImports & smWebHeaders
        
        'If Not gFTPFileFromWebServer(slPathFileName, smWebHeaders) Then
        
        
        gSleep (3)
        llCnt = 0
        If Not mFtpWebFileToServer(smWebHeaders) Then
            'potential for an endless loop
            While Not mFtpWebFileToServer(smWebHeaders)
            Wend
            
            Close #hmTo
            SetResults "No headers were found to import at the present time.", RGB(0, 0, 200)
            gLogMsg "No headers were found to import at the present time.", "WebImportLog.Txt", False
            Screen.MousePointer = vbDefault
            imImporting = False
            cmdCancel.Enabled = True
            Call gEndWebSession("WebImportLog.Txt")
            Exit Sub
        End If
    
        SetResults "Receiving spots from web site.", 0
        gLogMsg "Receiving spots from web site.", "WebImportLog.Txt", False
        slPathFileName = smWebImports & smWebSpots
        'If Not gFTPFileFromWebServer(slPathFileName, smWebSpots) Then
        If Not mFtpWebFileToServer(smWebSpots) Then
            Close #hmTo
            SetResults "No spots were found to import at the present time.", RGB(0, 0, 200)
            gLogMsg "No spots were found to import at the present time.", "WebImportLog.Txt", False
            Screen.MousePointer = vbDefault
            imImporting = False
            cmdCancel.Enabled = True
            Call gEndWebSession("WebImportLog.Txt")
            Exit Sub
        End If
    
        SetResults "Receiving activity logs from web site.", 0
        gLogMsg "Receiving activity logs from web site.", "WebImportLog.Txt", False
        slPathFileName = smWebImports & smWebLogs
        'If Not gFTPFileFromWebServer(slPathFileName, smWebLogs) Then
        If Not mFtpWebFileToServer(smWebLogs) Then
            'Dont let this cause the process to fail
            'Print #hmMsg, "** Terminated **"
            'Close #hmMsg
            'Close #hmTo
            'lacResult.Caption = "No data to import at the present time."
            'lacResult.ForeColor = RGB(0, 155, 0)
            'Screen.MousePointer = vbDefault
            'imExporting = False
            'cmdCancel.Enabled = True
            'Exit Sub
        End If
    
        Call gEndWebSession("WebImportLog.Txt")
    End If
    If Not mCheckFile() Then
        Screen.MousePointer = vbDefault
        SetResults "mCheckFile: No data was found to import at the present time.", RGB(255, 0, 0)
        gLogMsg "mCheckFile: No data was found to import at the present time.", "WebImportLog.Txt", False
        cmdCancel.Enabled = True
        imImporting = False
        Exit Sub
    End If

    SetResults "Importing headers.", 0
    gLogMsg "Importing headers.", "WebImportLog.Txt", False
    iRet = mImportHeaders()
    If (iRet = False) Then
        Close #hmTo
        SetResults "Importing headers failed.", RGB(255, 0, 0)
        gLogMsg "ERROR: IMPORT - Importing headers failed.", "WebImportLog.Txt", False
        lbcMsg.ForeColor = RGB(255, 0, 0)
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.Enabled = True
        cmdCancel.SetFocus
        Exit Sub
    End If
    SetResults Str(lmTotalHeaders) & " headers imported.", 0
    gLogMsg CStr(lmTotalHeaders) & " headers imported.", "WebImportLog.Txt", False
    SetResults "Importing spots.", 0
    gLogMsg "Importing spots.", "WebImportLog.Txt", False
    
    
    'D.S. 02/05/13 Added call below
    gEraseEventDate
    '7458
    Set myEnt = New CENThelper
    With myEnt
        .TypeEnt = Receivedpostedfromweb
        .User = igUstCode
        .ThirdParty = Web
        .ErrorLog = cmPathForgLogMsg
        .ProcessStart
    End With
    bgTaskBlocked = False
    sgTaskBlockedName = "Counterpoint Affidavit System Import"
    iRet = mProcessSpotFile()
    SetResults Str(lmTotalSpots) & " spots imported.", 0
    gLogMsg CStr(lmTotalSpots) & " spots imported.", "WebImportLog.Txt", False
    
    If (iRet = False) Then
        If lmTotalSpots = 0 Then
            SetResults "Requesting web site to commit changes.", 0
            gLogMsg "Requesting web site to commit changes.", "WebImportLog.Txt", False
            If Not gSendCmdToWebServer("ExportCommit.dll", Now()) Then
                SetResults "Error: Unable to instruct Web Server to commit changes.", RGB(255, 0, 0)
                gLogMsg "ERROR: IMPORT - Unable to instruct Web Server to commit changes.", "WebImportLog.Txt", False
                imImporting = False
            End If
        End If
    
        bgTaskBlocked = False
        sgTaskBlockedName = ""
        Close #hmTo
        SetResults "Importing spots failed.", RGB(255, 0, 0)
        gLogMsg "ERROR: IMPORT - Importing spots failed.", "WebImportLog.Txt", False
        imImporting = False
        Screen.MousePointer = vbDefault
        cmdCancel.Enabled = True
        cmdCancel.SetFocus
        Exit Sub
    End If

    SetResults "Importing activity logs.", 0
    gLogMsg "Importing activity logs.", "WebImportLog.Txt", False
    iRet = mImportLogs()
    SetResults Str(lmTotalActivityLogs) & " activity logs imported.", 0
    gLogMsg CStr(lmTotalActivityLogs) & " activity logs imported.", "WebImportLog.Txt", False

    If (StrComp(sgCommand, "/m", vbTextCompare) <> 0) Then
        SetResults "Requesting web site to commit changes.", 0
        gLogMsg "Requesting web site to commit changes.", "WebImportLog.Txt", False
        If Not gSendCmdToWebServer("ExportCommit.dll", Now()) Then
            bgTaskBlocked = False
            sgTaskBlockedName = ""
            Close #hmTo
            SetResults "Error: Unable to instruct Web Server to commit changes.", RGB(255, 0, 0)
            gLogMsg "ERROR: IMPORT - Unable to instruct Web Server to commit changes.", "WebImportLog.Txt", False
            Screen.MousePointer = vbDefault
            cmdCancel.Enabled = True
            imImporting = False
            Exit Sub
        End If
    End If
    
    If (StrComp(sgCommand, "/m", vbTextCompare) <> 0) Then
        'Erase spots if necessary
        'SetResults "Checking for Spots to Erase: " & smEarliestSpottDate, 0
        llRet = mTestToErase()
        If llRet = -1 Then
            'Error the function did not execute correctly.
            SetResults "FAIL: Erasing had Error(s).", RGB(255, 0, 0)
            SetResults "Erasing Did NOT Complete", RGB(255, 0, 0)
            gLogMsg "ERROR: IMPORT - Erasing had Error(s).", "WebImportLog.Txt", False
        End If
        
        SetResults "Erase Check Complete", 0
        If llRet > 0 Then
            SetResults CStr(llRet) & " Spots Required Erasing", 0
            gLogMsg CStr(llRet) & " Spots Required Erasing", "WebImportLog.Txt", False
        Else
            SetResults "None Required Erasing", 0
            gLogMsg "No Spots Required Erasing", "WebImportLog.Txt", False
        End If
    End If
    
    On Error GoTo ErrHand:
    
    If bgTaskBlocked And (Not igAutoImport) Then
         SetResults "Some spots were blocked during Import.", MESSAGERED
         gMsgBox "Some spots were blocked during the Import." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    
    SetResults "Import Completed Successfully.", RGB(0, 155, 0)
    
    lbcMsg.ListIndex = -1   ' Finish with nothing selected
    imImporting = False
    gLogMsg "** Completed Web Import Aired Station Spots. **", "WebImportLog.Txt", False
    gLogMsg "", "WebImportLog.Txt", False
    
    Screen.MousePointer = vbDefault
    
    'D.S. TTP 9859 Stop writing out Nothing.txt file that can cause permissions problems
    'If Not gExecExtStoredProc("Nothing.txt", "ImportFailedEmails.exe", True, True) Then
    If Not gExecExtStoredProc("Nothing.txt", "ImportFailedEmails.exe", False, False) Then
        SetResults "Unable to instruct Web site to Send Failed Emails...", RGB(255, 0, 0)
        gLogMsg "Error: " & "Unable to instruct Web site to Send Failed Emails", "WebImportLog.Txt", False
        gLogMsg "", "WebImportLog.Txt", False
        Screen.MousePointer = vbDefault
    End If
    
    lgETime5 = timeGetTime
    lgTtlTime5 = lgTtlTime5 + (lgETime5 - lgSTime5)
    
    mLogTimingResults
    
    DoEvents
    cmdCancel.Caption = "&Done"
    cmdImport.Enabled = True
    DoEvents
    '7458
    Set myEnt = Nothing
    Exit Sub

cmdImportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "cmdImport Click"
    Call gEndWebSession("WebImportLog.Txt")
    imImporting = False
    Exit Sub
End Sub


Private Function mImportHeaders() As Integer
    'On Error GoTo mImportHeadersErr_1
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim ilShfCode As Integer
    Dim SHTT_Recs As ADODB.Recordset
    Dim ATT_Recs As ADODB.Recordset

    mImportHeaders = False
    slFromFile = smWebImports & smWebHeaders
    
    'ilRet = 0
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read Lock Write As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read Lock Write", hmFrom)
    If ilRet <> 0 Then
        Exit Function
    End If

    Dim attCode, NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, ImportType, LogType, PostType, startTime, StationEmail, StationPW, AggreementEmail, AggreementPW As String
    ' Skip past the header definition record.
    Input #hmFrom, attCode, NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, ImportType, LogType, PostType, startTime, StationEmail, StationPW, AggreementEmail, AggreementPW
    ' Verify we are indeed processing a WebHeaders.txt file.
    If Len(attCode) < 1 Or attCode <> "attcode" Then
        Exit Function
    End If

    On Error GoTo mImportHeadersErr_2
    Do While Not EOF(hmFrom)
        lmTotalHeaders = lmTotalHeaders + 1
        Input #hmFrom, attCode, NetworkSWProvider, WebsiteProvider, StationProvider, NetworkName, VehicleName, StationName, ImportType, LogType, PostType, startTime, StationEmail, StationPW, AggreementEmail, AggreementPW
        If attCode > 0 Then
            SQLQuery = "Select attWebPW, attshfCode From ATT Where attCode = " & attCode
            Set ATT_Recs = gSQLSelectCall(SQLQuery)
            If Not ATT_Recs.EOF Then
                ilShfCode = ATT_Recs!attshfcode
                SQLQuery = "Update ATT Set attWebPW = '" & Trim(AggreementPW) & "', attWebEmail = '" & Trim(AggreementEmail) & "' Where attCode = " & attCode
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/13/16: Replaced GoSub
                    'GoSub mImportHeadersErr_2:
                    Screen.MousePointer = vbDefault
                    gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportHeaders"
                    mImportHeaders = False
                    Close hmFrom
                    Exit Function
                End If
                ' Compare the shttWebPW too
                SQLQuery = "Select shttWebPW From SHTT Where shttCode = " & ilShfCode
                Set SHTT_Recs = gSQLSelectCall(SQLQuery)
                If Not SHTT_Recs.EOF Then
                    SQLQuery = "Update SHTT Set shttWebPW = '" & Trim(StationPW) & "', shttWebEmail = '" & Trim(StationEmail) & "' Where shttCode = " & ilShfCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub mImportHeadersErr_2:
                        Screen.MousePointer = vbDefault
                        gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportHeaders"
                        mImportHeaders = False
                        Close hmFrom
                        Exit Function
                    End If
                    '11/26/17
                    mUpdateShttTables ilShfCode, Trim$(StationPW), Trim(StationEmail)
                End If
                SHTT_Recs.Close
            End If
            ATT_Recs.Close
        Else
            SQLQuery = "Select latShttCode from lat where latWebLogAttID = " & Abs(attCode)
            Set lat_recs = gSQLSelectCall(SQLQuery)
            If Not lat_recs.EOF Then
                ilShfCode = lat_recs!latShttCode
                ' Compare the shttWebPW too
                SQLQuery = "Select shttWebPW From SHTT Where shttCode = " & ilShfCode
                Set SHTT_Recs = gSQLSelectCall(SQLQuery)
                If Not SHTT_Recs.EOF Then
                    SQLQuery = "Update SHTT Set shttWebPW = '" & Trim(StationPW) & "', shttWebEmail = '" & Trim(StationEmail) & "' Where shttCode = " & ilShfCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub mImportHeadersErr_2:
                        Screen.MousePointer = vbDefault
                        gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportHeaders"
                        mImportHeaders = False
                        Close hmFrom
                        Exit Function
                    End If
                    '11/26/17
                    mUpdateShttTables ilShfCode, Trim$(StationPW), Trim(StationEmail)
                End If
                SHTT_Recs.Close
            End If
        End If
    Loop
    mImportHeaders = True
    Close hmFrom
    Exit Function

'mImportHeadersErr_1:
'    ilRet = Err.Number
'    Resume Next

mImportHeadersErr_2:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mImportHeaders"
    Exit Function
End Function


Private Function mImportLogs() As Integer
    'On Error GoTo mImportLogsErr_1
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim SHTT_Recs As ADODB.Recordset
    Dim ATT_Recs As ADODB.Recordset

    mImportLogs = False
    slFromFile = smWebImports & smWebLogs
    
    'ilRet = 0
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read Lock Write As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read Lock Write", hmFrom)
    If ilRet <> 0 Then
        Exit Function
    End If

    ' Skip past the header definition record.
    Input #hmFrom, tmweblCode, tmweblType, tmweblattCode, tmweblIPAddr, tmweblPCName, tmweblCallLetters, tmweblVehicleName, tmwebUserName, tmwebPostDay, tmweblDate, tmWeblTime
    ' Verify we are indeed processing a WebLogs.txt file.
    If Len(tmweblCode) < 1 Or tmweblCode <> "weblCode" Then
        Close hmFrom
        Exit Function
    End If

    On Error GoTo mImportLogsErr_2
    Do While Not EOF(hmFrom)
        lmTotalActivityLogs = lmTotalActivityLogs + 1
        Input #hmFrom, tmweblCode, tmweblType, tmweblattCode, tmweblIPAddr, tmweblPCName, tmweblCallLetters, tmweblVehicleName, tmwebUserName, tmwebPostDay, tmweblDate, tmWeblTime
        SQLQuery = "Insert Into WebL (weblType, weblattCode, weblIP, weblCPUName, weblCallLetters, weblVehicleName, weblUserName, weblPostDay, weblDate, weblTime) "
        tmweblVehicleName = gFixQuote(tmweblVehicleName)
        tmwebUserName = gFixQuote(tmwebUserName)
        tmweblPCName = gFixQuote(tmweblPCName)
        SQLQuery = SQLQuery & "Values ("
        SQLQuery = SQLQuery & tmweblType & ", "
        SQLQuery = SQLQuery & tmweblattCode & ", "
        SQLQuery = SQLQuery & "'" & tmweblIPAddr & "', "
        SQLQuery = SQLQuery & "'" & tmweblPCName & "', "
        SQLQuery = SQLQuery & "'" & tmweblCallLetters & "', "
        SQLQuery = SQLQuery & "'" & tmweblVehicleName & "', "
        SQLQuery = SQLQuery & "'" & tmwebUserName & "', "
        SQLQuery = SQLQuery & "'" & Format$(tmwebPostDay, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "'" & Format$(tmweblDate, sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "'" & Format$(tmWeblTime, sgSQLTimeForm) & "'"
        SQLQuery = SQLQuery & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub mImportHeadersErr_2:
            Screen.MousePointer = vbDefault
            gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportHeaders"
            mImportLogs = False
            Close hmFrom
            Exit Function
        End If
    Loop
    Close hmFrom
    mImportLogs = True
    Exit Function

mImportLogsErr_1:
    ilRet = Err.Number
    Resume Next

mImportLogsErr_2:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mImportLogs"
    Close hmFrom
    Exit Function
End Function


Private Function mProcessSpotFile() As Integer

    'Created by D.S. June 2007
    'Process the import file sent back by the web server
    
    Dim slFromFile As String
    Dim slLine As String
    Dim slSDate As String
    Dim slFeedDate As String
    Dim slFeedTime As String
    Dim slSunDate As String
    Dim slMonDate As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llIdx As Long
    Dim llTotalSpots As Long
    Dim llAstCode As Long
    Dim llAttCode As Long
    Dim ilFirstTime As Integer
    Dim slNoAstExists As String
    Dim ilWriteBlank As Integer
    Dim cprst As ADODB.Recordset
    Dim ilVefCode As Integer
    Dim ilAdfCode As Integer
    Dim ilAst As Integer
    Dim ilAnyAstExist As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim ilAnyNotCompliant As Integer
    Dim ilRetryCount As Integer
    Dim slTemp As String
    Dim ilFillStatus As Integer
    '8862
    Dim ilVendorCodesToUpdate() As Integer
    
    On Error GoTo ErrHand
    mProcessSpotFile = False
    lmResolvedBytime = 0
    lmTotalSpots = 0
    llIdx = 0
    slFromFile = smWebImports & smWebSpots
    
    'debug test file
    'slFromFile = "D:\CSI\V55\import\WebLogs_WebSpots_DG.txt"
    'slFromFile = "D:\CSI\V55\import\Web_Spot_Test.txt"
    
    'ilRet = 0
    'On Error GoTo mImportSpotsErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read Lock Write As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read Lock Write", hmFrom)
    If ilRet <> 0 Then
        sgReImportStatus = "Counterpoint Affidavit: Unable to Open File " & slFromFile
        Close hmFrom
        Exit Function
    End If
    Line Input #hmFrom, slLine  ' Skip past the header definition record.
    ReDim tmAirSpotInfo(0 To 0) As AIRSPOTINFO
    ReDim tmTempAirSpotInfo(0 To 0) As AIRSPOTINFO
    
    llTotalSpots = 0
    '7458
    myEnt.fileName = smWebSpots
    Do While Not EOF(hmFrom)
        DoEvents
        'Process Input
        On Error GoTo ErrHand
        DoEvents
        imNoSpotsAired = True
        '8/3/11: Added test for which pass
        'ilRet = mFillAirSpotInfo(llIdx, True)
        If llIdx = 0 Then
            ilFillStatus = mFillAirSpotInfo(llIdx, True)
            If ilFillStatus <> 1 Then
                Exit Do
            End If
        Else
            llIdx = 0
        End If
        llAttCode = tmAirSpotInfo(0).lAtfCode
        ilVefCode = gGetVehCodeFromAttCode(CStr(llAttCode))
        imVefCode = ilVefCode
        smPledgeByEvent = mGetPledgeByEvent()
        llAstCode = Val(tmAirSpotInfo(0).lAstCode)
        slFeedDate = Format(tmAirSpotInfo(0).sFeedDate, sgSQLDateForm)
        slFeedTime = Format(tmAirSpotInfo(0).sFeedTime, sgSQLTimeForm)
        slMonDate = gAdjYear(gObtainPrevMonday(slFeedDate))
        slSunDate = gObtainNextSunday(slMonDate)
        imEOF = False
        'Do Until EOF(hmFrom) Or (DateValue(tmAirSpotInfo(llIdx).sFeedDate) < DateValue(slMonDate)) Or (DateValue(tmAirSpotInfo(llIdx).sFeedDate) > DateValue(slSunDate))
        'Do Until EOF(hmFrom) Or tmAirSpotInfo(llIdx).lAtfCode <> llAttCode Or (DateValue(tmAirSpotInfo(llIdx).sFeedDate) < DateValue(slMonDate)) Or (DateValue(tmAirSpotInfo(llIdx).sFeedDate) > DateValue(slSunDate))
        '8/3/11

        'D.S. 6/11/12 new code reDim
        ReDim Preserve tmAirSpotInfo(0 To 500) As AIRSPOTINFO
        '8862 store the vendor ids we find in mFillAirSpotInfo (as 'source') that are allowed to auto update
        ReDim ilVendorCodesToUpdate(0 To 0) As Integer
        Do Until EOF(hmFrom)
            If tmAirSpotInfo(llIdx).lAtfCode <> llAttCode Or (DateValue(tmAirSpotInfo(llIdx).sFeedDate) < DateValue(slMonDate)) Or (DateValue(tmAirSpotInfo(llIdx).sFeedDate) > DateValue(slSunDate)) Then
                Exit Do
            Else
                llIdx = llIdx + 1
                'D.S. 6/11/12 new code if statement
                If llIdx > UBound(tmAirSpotInfo) Then
                    ReDim Preserve tmAirSpotInfo(0 To llIdx + 250) As AIRSPOTINFO
                End If
                ilFillStatus = mFillAirSpotInfo(llIdx, True)
            End If
        Loop
        ReDim Preserve tmAirSpotInfo(0 To llIdx) As AIRSPOTINFO
        If EOF(hmFrom) Then
            If ilFillStatus = 1 Then
            llIdx = llIdx + 1
            End If
            imEOF = True
        End If
        
        ReDim Preserve tmAirSpotInfo(0 To llIdx) As AIRSPOTINFO
        'D.S. 10/3/10 Added call to gGetAstInfo
        'This assures that we have the latest ast/spots. It covers the case where spots were exported to the web
        'and then spots were changed in Traffic, but never exported to the web after the logs were generated.
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attAgreeStart, attAgreeEnd, attDropDate, attOnAir, attOffAir, attMulticast "
        SQLQuery = SQLQuery & " FROM shtt, cptt, att "
'        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attAgreeStart, attAgreeEnd, attDropDate, attOnAir, attOffAir, attMulticast, attWebInterface, attExportToUnivision, attExportToMarketron, attExportToCBS, attExportToClearCh, attExportToJelli, attAudioDelivery"
'        SQLQuery = SQLQuery & " FROM shtt, cptt, att"
        SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
        SQLQuery = SQLQuery & " AND attExportType = 1"
        'D.S. 11/27/13
        'SQLQuery = SQLQuery & " AND cpttVefCode = " & ilVefCode
        SQLQuery = SQLQuery & " AND cpttAtfCode = " & llAttCode
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slMonDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " Order by shttcallletters"
        Set cprst = gSQLSelectCall(SQLQuery)
        If Not cprst.EOF Then
            'Create AST records - gGetAstInfo requires tgCPPosting to be initialized
            ReDim tgCPPosting(0 To 1) As CPPOSTING
            tgCPPosting(0).lCpttCode = cprst!cpttCode
            tgCPPosting(0).iStatus = cprst!cpttStatus
            tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
            tgCPPosting(0).lAttCode = cprst!cpttatfCode
            tgCPPosting(0).iAttTimeType = cprst!attTimeType
            tgCPPosting(0).iVefCode = ilVefCode
            tgCPPosting(0).iShttCode = cprst!shttCode
            tgCPPosting(0).sZone = cprst!shttTimeZone
            tgCPPosting(0).sDate = Format$(slMonDate, sgShowDateForm)
            tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
            igTimes = 1 'By Week
            DoEvents
            '7458
            With myEnt
                .Vehicle = ilVefCode
                .Station = cprst!shttCode
                .Agreement = tmAirSpotInfo(0).lAtfCode
                If Not .SetThirdPartyByHierarchy() Then
                    gLogMsg .ErrorMessage, "WebImportLog.Txt", False
                End If
                .ProcessStart
            End With
            'D.S. 02/05/19
            ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), -1, True, True, True) ' , False)

            If Not ilRet Then
                ilRet = ilRet
            End If
        Else
            ReDim tmAstInfo(0 To 0)
            ilRet = ilRet
        End If
        'D.S. End new code
        
        'D.S. 11/4/15 loop through tmAstInfo and set all iCpstatus = 0
        For ilLoop = 0 To UBound(tmAstInfo) - 1 Step 1
            tmAstInfo(ilLoop).iCPStatus = 0
        Next ilLoop
        
        ilRet = mImportByAstCode(slSunDate, slMonDate)
        If imNextPassNeeded Then
            ilRet = mImportByFeedDate(slSunDate, slMonDate)
            If imImportByAstCodeAgain Then
                'we found at least one spot based off of the FeedDate so run it back through the
                'first pass so it can be updated
                ilRet = mImportByAstCode(slSunDate, slMonDate)
            End If
        End If
        

        ilRet = False
        ilRetryCount = 0
        While Not ilRet And ilRetryCount < 5
            ilRet = mUpdateCptt()
            ilRetryCount = ilRetryCount + 1
            Sleep 250  'Delay 1 second before retrying
        Wend
        '8862
        mAddToVendorUpdateAsNeeded ilVendorCodesToUpdate()
        If Not gUpdateVendorStatusAsNeeded(llAttCode, ilVendorCodesToUpdate) Then
            SetResults "warning!  Issue with Auto Vendor Delivery Status. Please see log", MESSAGERED
           ' myErrors.WriteError "Issue with Auto Vendor Delivery Status. gUpdateVendorStatusAsNeeded-" & Err.Description & "  See 'AffErrorLog.txt' for more information.", False
            gLogMsg "Issue with Auto Vendor Delivery Status. gUpdateVendorStatusAsNeeded-" & Err.Description & "  See 'AffErrorLog.txt' for more information.", cmPathForgLogMsg, False
        End If
        If Not ilRet Then
            '7458
            myEnt.ClearWhenDontSend
            sgReImportStatus = "Counterpoint Affidavit: Failed, see WebImportLog.Txt for error"
            Close hmFrom
            Exit Function
        End If
        '8/14/18: Removed as last agreement compliant info was not being set
        ''D.S. 05-04-17 Moved down from above
        'If EOF(hmFrom) Then
        '    Exit Do
        'End If
        
        'D.S. 02/25/11 Start new compliant code
        SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attCode, attTimeType, attGenCP, attStartTime, attLogType, attPostType, shttWebEmail, shttWebPW, attWebEmail, attSendLogEmail, attWebPW, attAgreeStart, attAgreeEnd, attDropDate, attOnAir, attOffAir, attMulticast "
        SQLQuery = SQLQuery & " FROM shtt, cptt, att"
        SQLQuery = SQLQuery & " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery & " AND attCode = cpttAtfCode"
        SQLQuery = SQLQuery & " AND attExportType = 1"
        'D.S. 11/27/13
        'SQLQuery = SQLQuery & " AND cpttVefCode = " & ilVefCode
        SQLQuery = SQLQuery & " AND cpttAtfCode = " & llAttCode
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slMonDate, sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " Order by shttcallletters"
        Set cprst = gSQLSelectCall(SQLQuery)
        If Not cprst.EOF Then
            'Create AST records - gGetAstInfo requires tgCPPosting to be initialized
            ReDim tgCPPosting(0 To 1) As CPPOSTING
            tgCPPosting(0).lCpttCode = cprst!cpttCode
            tgCPPosting(0).iStatus = cprst!cpttStatus
            tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
            tgCPPosting(0).lAttCode = cprst!cpttatfCode
            tgCPPosting(0).iAttTimeType = cprst!attTimeType
            tgCPPosting(0).iVefCode = ilVefCode
            tgCPPosting(0).iShttCode = cprst!shttCode
            tgCPPosting(0).sZone = cprst!shttTimeZone
            tgCPPosting(0).sDate = Format$(slMonDate, sgShowDateForm)
            tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
            ilSchdCount = 0
            ilAiredCount = 0
            ilPledgeCompliantCount = 0
            ilAgyCompliantCount = 0
            igTimes = 1 'By Week
               
            ilAdfCode = -1
            'Dan M 9/26/13  6442 changed to as v60, per Dick
            'D.S. 02/05/19
            ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, True, True, True) ', False)
            'ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, False, False, True, False)
            For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                ilAnyAstExist = True
                mCheckForMG tmAstInfo(ilAst)
                'gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, tmAstInfo(ilAst).iStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                gIncSpotCounts tmAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
                '7458
                If Not myEnt.Add(tmAstInfo(ilAst).sFeedDate, tmAstInfo(ilAst).lgsfCode, Asts) Then
                    gLogMsg myEnt.ErrorMessage, "WebImportLog.Txt", False
                End If
            Next ilAst
            If ilAiredCount <> ilPledgeCompliantCount Then
                ilAnyNotCompliant = True
            End If
            SQLQuery = "Update cptt Set "
            SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
            SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
            SQLQuery = SQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
            SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "cpttAgyCompliant = " & ilAgyCompliantCount & " "
            SQLQuery = SQLQuery & " Where cpttCode = " & cprst!cpttCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "WebImportLog.txt", "WebImportSchdSpot-mProcessSpotFile"
                mProcessSpotFile = False
                Close hmFrom
                Exit Function
            End If
        End If
        'D.S. 02/25/11 End new code
        ilWriteBlank = False
        For llIdx = 0 To UBound(tmAirSpotInfo) - 1 Step 1
            If tmAirSpotInfo(llIdx).iFound = False Then
                slNoAstExists = "Warning: " & tmAirSpotInfo(llIdx).lAtfCode & ","
                slTemp = gFindVehStaFromAttCode(CStr(tmAirSpotInfo(llIdx).lAtfCode))
                If slTemp = "Unable to find Agreement Code " Then
                    slNoAstExists = slNoAstExists & "," & ","
                Else
                    slNoAstExists = slNoAstExists & slTemp
                End If
                slNoAstExists = slNoAstExists & Trim$(gStripComma(tmAirSpotInfo(llIdx).sAdvt)) & ","
                slNoAstExists = slNoAstExists & Trim$(gStripComma(tmAirSpotInfo(llIdx).sProd)) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sPledgeStartDate1) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sPledgeEndDate) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sPledgeStartTime) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sPledgeEndTime) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).iSpotLen) & ","
                slNoAstExists = slNoAstExists & Trim$(gStripComma(tmAirSpotInfo(llIdx).sCart)) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sISCI) & ","
                slNoAstExists = slNoAstExists & Trim$(gStripComma(tmAirSpotInfo(llIdx).sCreativeTitle)) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).lAstCode) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sActualAirDate1) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sActualAirTime1) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sStatusCode) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sFeedDate) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sFeedTime) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sEndDate) & ","
                slNoAstExists = slNoAstExists & Trim$(tmAirSpotInfo(llIdx).sStartDate) & ","
                DoEvents
                If slTemp = "Unable to find Agreement Code " Then
                    gLogMsg slNoAstExists & "Unable to find Agreement Code ", "WebImportLog.Txt", False
                    'D.S. 5/1/15 No need to dosplay all of the to the screen. The list box can on handle 32,767 entries and is prone to stack overflow
                    'SetResults "Unable to process: " & slNoAstExists & slTemp, RGB(255, 0, 0)
                Else
                    If sgCommand <> "/ReImport" Then
                        gLogMsg slNoAstExists & " AST missing and No Time Match", "WebImportLog.Txt", False
                    End If
                    'D.S. 5/1/15 No need to display all of the to the screen. The list box can on handle 32,767 entries and is prone to stack overflow
                    'SetResults "Unable to process: " & slNoAstExists & slTemp, RGB(255, 0, 0)
                End If
                
                'SetResults "Unable to process: " & slNoAstExists & " AST missing or No Time Match", RGB(255, 0, 0)
                'gLogMsg "Warning: " & slNoAstExists & " AST missing and No Time Match", "WebImportLog.Txt", False
                ilWriteBlank = True
                '7458
                If Not myEnt.Add(tmAirSpotInfo(llIdx).sFeedDate, 0) Then
                        gLogMsg myEnt.ErrorMessage, "WebImportLog.Txt", False
                End If
            End If
        Next llIdx
        'If ilWriteBlank Then
        '    gLogMsg "", "WebImportLog.Txt", False
        'End If
        
        '7458
        If Not myEnt.CreateEnts() Then
            gLogMsg myEnt.ErrorMessage, "WebImportLog.Txt", False
        End If
        
        '8/3/11: Exit look if all records processed
        If imEOF Then
            Exit Do
        End If
        'prepare for a new week
        'ReDim tmAirSpotInfo(0 To llIdx) As AIRSPOTINFO
        llIdx = UBound(tmAirSpotInfo)
        'LSet tmTempAirSpotInfo(0) = tmAirSpotInfo(llIdx)
        tmTempAirSpotInfo(0) = tmAirSpotInfo(llIdx)
        ReDim tmAirSpotInfo(0 To 0) As AIRSPOTINFO
        llIdx = 0
        'LSet tmAirSpotInfo(llIdx) = tmTempAirSpotInfo(0)
        tmAirSpotInfo(llIdx) = tmTempAirSpotInfo(0)
        'Fill in the the first record with the info that's already in the slFields buffer.  This is due
        'do the fact that we went one too many records above
        'ilRet = mFillAirSpotInfo(llIdx, False)
        llIdx = llIdx + 1
        ReDim Preserve tmAirSpotInfo(0 To llIdx) As AIRSPOTINFO
    Loop
    
    mProcessSpotFile = True
    If sgCommand <> "/ReImport" Then
        gLogMsg "**** Resolved by Time **** " & CStr(lmResolvedBytime), "WebImportLog.Txt", False
    End If
    Close hmFrom
    Exit Function
    
'mImportSpotsErr:
'    ilRet = Err.Number
'    Resume Next
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mProcessSpotFile"
    mProcessSpotFile = False
    sgReImportStatus = "Counterpoint Affidavit: Failed, see WebImportLog.Txt for error"
    Close hmFrom
    Exit Function
End Function

Private Function mImportByAstCode(sSunDate As String, sMonDate As String) As Integer

    'Created by D.S. June 2007
    'Try to update the AST by looking at the AST code.  If it exists then update the record and
    'mark it as found in the array.  Otherwise mark it not found.
    
    Dim llLoop As Long
    Dim llAstCode As Long
    Dim ilAstFound As Integer
    Dim rst As ADODB.Recordset
    Dim rstMG As ADODB.Recordset
    Dim slSpotType As String
    Dim slInDate As String
    Dim slInTime As String
    Dim ilPos As Integer
    Dim slNoAstExists As String
    Dim slLine As String
    Dim ilRet As Integer
    Dim ilIdx As Integer
    Dim llPrevGsfCode As Long
    Dim slClearStatus As String
    Dim ilLoop As Integer
    Dim ilAstStatus As Integer
    Dim slRecType As String
    Dim llOldCpfCode As Long
    Dim llNewCpfCode As Long
    Dim ilAdfCode As Integer
    Dim ilISCIChg As Integer
    Dim slSource As String
    Dim slTemp As String
    
    On Error GoTo ErrHand
    
    mImportByAstCode = False
    imNextPassNeeded = False
    llPrevGsfCode = -1
    For llLoop = LBound(tmAirSpotInfo) To UBound(tmAirSpotInfo) - 1 Step 1
        DoEvents
        'By having the UpdateComplete flag we can call this function multiple times and only process what needs processing
        If tmAirSpotInfo(llLoop).iUpdateComplete = False Then
            ilAstFound = False
            llAstCode = tmAirSpotInfo(llLoop).lAstCode
            slSource = tmAirSpotInfo(llLoop).sVendorSource
            If Trim(tmAirSpotInfo(llLoop).sRecType) = "M" Then
                imNoSpotsAired = False
                ilRet = mImportMG(llLoop)
                '7458
                If Not myEnt.Add(tmAirSpotInfo(llLoop).sFeedDate, tmAirSpotInfo(llLoop).lgsfCode, MakeGood) Then 'change Bonus to Replacement or MakeGood  as needed
                    gLogMsg myEnt.ErrorMessage, "WebImportLog.Txt", False
                End If
            ElseIf Trim(tmAirSpotInfo(llLoop).sRecType) = "B" Then
                imNoSpotsAired = False
                ilRet = mImportBonus(llLoop)
                '7458
                If Not myEnt.Add(tmAirSpotInfo(llLoop).sFeedDate, tmAirSpotInfo(llLoop).lgsfCode, Bonus) Then 'change Bonus to Replacement or MakeGood  as needed
                    gLogMsg myEnt.ErrorMessage, "WebImportLog.Txt", False
                End If
            ElseIf Trim(tmAirSpotInfo(llLoop).sRecType) = "R" Then
                imNoSpotsAired = False
                ilRet = mImportReplacement(llLoop)
                '7458
                If Not myEnt.Add(tmAirSpotInfo(llLoop).sFeedDate, tmAirSpotInfo(llLoop).lgsfCode, Replacement) Then 'change Bonus to Replacement or MakeGood  as needed
                    gLogMsg myEnt.ErrorMessage, "WebImportLog.Txt", False
                End If
            Else
                'Debug only
                'If llAstCode = 30909251 Then
                '    ilRet = ilRet
                'End If
                If smPledgeByEvent = "Y" Then
                    SQLQuery = "Select astCode, astAdfCode, astShfCode, astAtfCode, astVefCode, astFeedDate, astStatus, astCpfCode, lstGsfCode FROM ast left outer join lst on astLsfCode = lstCode WHERE (astCode = " & tmAirSpotInfo(llLoop).lAstCode & ")"
                Else
                    SQLQuery = "Select astCode, astAdfCode, astShfCode, astAtfCode, astVefCode, astFeedDate, astStatus, astCpfCode FROM ast WHERE (astCode = " & tmAirSpotInfo(llLoop).lAstCode & ")"
                End If
                Set rst = cnn.Execute(SQLQuery)
                
                If Not rst.EOF Then
                    tmAirSpotInfo(llLoop).iFound = False
                    If (rst!astShfCode = tmAirSpotInfo(llLoop).iShfCode) And (rst!astAtfCode = tmAirSpotInfo(llLoop).lAtfCode) And (rst!astVefCode = tmAirSpotInfo(llLoop).iVefCode) Then
                        If DateValue(gAdjYear(rst!astFeedDate)) < DateValue(gAdjYear(sSunDate)) Then
                            tmAirSpotInfo(llLoop).sStartDate = rst!astFeedDate
                        End If
                        If DateValue(gAdjYear(rst!astFeedDate)) > DateValue(gAdjYear(sMonDate)) Then
                            tmAirSpotInfo(llLoop).sEndDate = rst!astFeedDate
                        End If
                        tmAirSpotInfo(llLoop).iFound = True
                    End If
                    
                    If tmAirSpotInfo(llLoop).iFound Then
                    
                        'slFields(15) = 0 then the spot was received and it aired
                        
                        'D.S. Please let me know if this mapping can made more confusing. I don't think so!
                        
                        'Also if they pledge a 2 Delay B'cast or a 10 'Delay Comm/Prgm  it maps to a
                        '2 Delay B'cast per Jim F. 7/28/06
                        
                        'C - Program and Commercial Aired Live.
                        '   Web returns = 0  Status = 0  Screen = 1-Aired Live
                        
                        'N - Neither the spot nor the program aired.
                        '   Web returns = 1  Status = 4   Screen = 5-Not Aired Other
                        'D - Program and Commercial were both delayed.
                        '   Web returns = 2  Status = 9   Screen = 10-Delay Cmml/Prg
                        'S - Program did not air, but spot aired, either live or delayed.
                        '   Web returns = 3  Status = 10  Screen = 11-Air Cmml Only
                        'P - Program aired spot did not.
                        'Old rule
                        '   Web returns = 4  Status = 8   Screen = 9-Not Carried or Aired.
                        'New rule per Jim
                        '   Web returns = 4  Status = 4   Screen = 5-Not Aired Other
                        'K - Delay B'cast
                        '   Web returns = 5  Status = 1   Screen = 2-Delay B'cast
                        'G - Games not carried
                        '   Web returns = 6  Status = 8   Screen = 9-Not Carried or Aired.
                        'X - Neither the spot nor the program aired and the spot will NOT be made good.
                        '   Web returns = 11  Status = 14   Screen = 5-Not Aired Other
                                
                        'D.S. 06/20/17
                        'ilAstStatus = rst!astStatus
                        ilAstStatus = tmAirSpotInfo(llLoop).sStatusCode
                        Select Case Val(tmAirSpotInfo(llLoop).sStatusCode)
                            Case 0
                                slSpotType = "C" 'Program and Commercial Aired Live.
                                imNoSpotsAired = False
                            Case 1
                                slSpotType = "N" 'Neither the spot nor the program aired.
                            Case 2
                                slSpotType = "D" 'Program and Commercial were both delayed.
                                imNoSpotsAired = False
                            Case 3
                                slSpotType = "S" 'Program did not air, but spot aired, either live or delayed.
                                imNoSpotsAired = False
                            Case 4
                                slSpotType = "P" 'Not Carried or Aired.
                                '6/29/12: P is spot not aired status, therefore, don't set the flag as aired
                                'imNoSpotsAired = False
                            Case 5
                                slSpotType = "K" 'Delay B'cast
                                imNoSpotsAired = False
                            Case 6
                                slSpotType = "G" 'Game Not Carried
                                imNoSpotsAired = True
                            Case 11
                                slSpotType = "X" 'Neither the spot nor the program aired.
                                imNoSpotsAired = True
                        End Select
            
                        If smPledgeByEvent = "Y" Then
                            If llPrevGsfCode <> rst!lstGsfCode Then
                                If slSpotType = "G" Then
                                    slClearStatus = "N"
                                Else
                                    slClearStatus = "Y"
                                End If
                                SQLQuery = "Update pet Set "
                                SQLQuery = SQLQuery & "petClearStatus = '" & slClearStatus & "'"
                                SQLQuery = SQLQuery & " WHERE petAttCode = " & rst!astAtfCode & " And petGsfCode = " & rst!lstGsfCode
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/13/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportByASTCode"
                                    mImportByAstCode = False
                                    Exit Function
                                End If
                                llPrevGsfCode = rst!lstGsfCode
                            End If
                        End If
            
                        'D.S. 06/20/17
                        'ilAstStatus = rst!astStatus
                        
                        'Loop thru tmAirSpotInfo looking for matching AST code, if found then set .iCpStatus = 1
                        'For ilLoop = 0 To UBound(tmAirSpotInfo) - 1 Step 1
                        'D.S. 10/24/16 pointed for loop at the correct type def
                        'D.S. 06/20/17
                        For ilLoop = 0 To UBound(tmAstInfo) - 1 Step 1
                            If tmAstInfo(ilLoop).lCode = llAstCode Then
                                If tmAstInfo(ilLoop).iStatus = 4 Then
                                    'It's a MG so find the link code
                                    SQLQuery = "Select astCode,astLkAstCode  FROM ast WHERE (astCode = " & tmAirSpotInfo(llLoop).lAstCode & ")"
                                    Set rstMG = cnn.Execute(SQLQuery)
                                    If Not rstMG.EOF Then
                                        slTemp = rstMG!astLkAstCode
                                    End If
                                    'Delete the AST that was the MG for the original
                                    SQLQuery = "DELETE FROM Ast WHERE astCode = " & slTemp
                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                        GoSub ErrHand:
                                End If
                                    'update the original AST
                                SQLQuery = "Update AST Set "
                                SQLQuery = SQLQuery & "astLkAstCode = '" & 0 & "', "
                                SQLQuery = SQLQuery & "astMissedMnfCode = '" & 0 & "'"
                                SQLQuery = SQLQuery & " WHERE astCode = " & tmAirSpotInfo(llLoop).lAstCode
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    GoSub ErrHand:
                                End If
                                If sgCommand = "/ReImport" Then
                                    ilAstStatus = tmAstInfo(ilLoop).iPledgeStatus
                                End If
                            Exit For
                        End If
                            End If
                        Next ilLoop
            
'                        D.S. 06/20/17 Removed below and replaced with the above For Loop.
'                        For ilLoop = 0 To UBound(tmAstInfo) - 1 Step 1
'                            If tmAstInfo(ilLoop).lCode = llAstCode Then
'                                tmAstInfo(ilLoop).iCPStatus = 1
'                                If sgCommand = "/ReImport" Then
'                                    ilAstStatus = tmAstInfo(ilLoop).iPledgeStatus
'                                End If
'                                Exit For
'                            End If
'                        Next ilLoop
            
            
            
                        If slSpotType = "C" Or slSpotType = "D" Or slSpotType = "S" Or slSpotType = "K" Then
                            'update date/time aired
                            llOldCpfCode = rst!astCpfCode
                            'llNewCpfCode = llOldCpfCode
                            ilAdfCode = rst!astAdfCode
                            llNewCpfCode = mGetCpfCode(ilAdfCode, Trim(tmAirSpotInfo(llLoop).sISCI))
        
                            'D.S. 02/5/19 start new code
                            tmAirSpotInfo(llLoop).bIsciChngFlag = True
                            If llNewCpfCode = llOldCpfCode Then
                                tmAirSpotInfo(llLoop).bIsciChngFlag = False
                            End If
                            'D.S. 02/5/19 end new code
                            
                            If tmAirSpotInfo(llLoop).bIsciChngFlag Then
                                llNewCpfCode = mAdjustISCIAsNeeded(tmAirSpotInfo(llLoop).lAstCode, ilAdfCode, llOldCpfCode, tmAirSpotInfo(llLoop).sISCI, "C")
                            End If
                            ilISCIChg = 0
                            If llNewCpfCode <> llOldCpfCode Then
                                ilISCIChg = ASTEXTENDED_ISCICHGD
                            End If
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            'D.S. 10/26/19 below astLen line
                            SQLQuery = SQLQuery + "astLen = " & tmAirSpotInfo(llLoop).iSpotLen & ", "
                            If slSpotType = "C" Then
                                'If gGetAirStatus(rst!astStatus) <= 1 Or gGetAirStatus(rst!astStatus) = 9 Or gGetAirStatus(rst!astStatus) = 10 Then
                                '    SQLQuery = SQLQuery + "astStatus = " & gGetAirStatus(rst!astStatus) & ", "  'Aired Live
                                If gGetAirStatus(ilAstStatus) <= 1 Or gGetAirStatus(ilAstStatus) = 9 Or gGetAirStatus(ilAstStatus) = 10 Then
                                    SQLQuery = SQLQuery + "astStatus = " & gGetAirStatus(ilAstStatus) + ilISCIChg & ", "  'Aired Live
                                Else
                                    SQLQuery = SQLQuery + "astStatus = " & 1 + ilISCIChg & ", " 'Delay B'cast
                                End If
                            End If
                            If slSpotType = "D" Then
                                SQLQuery = SQLQuery + "astStatus = " & 9 + ilISCIChg & ", " 'Program and Commercial were both delayed.
                            End If
                            If slSpotType = "S" Then
                                SQLQuery = SQLQuery + "astStatus = " & 10 + ilISCIChg & ", " 'Program did not air, but spot aired, either live or delayed.
                            End If
                            If slSpotType = "K" Then
                                SQLQuery = SQLQuery + "astStatus = " & 1 + ilISCIChg & ", " 'Delay B'cast
                            End If
                            SQLQuery = SQLQuery + "astCpfCode = " & llNewCpfCode & ", "
                            slInDate = tmAirSpotInfo(llLoop).sActualAirDate1
                            slInTime = tmAirSpotInfo(llLoop).sActualAirTime1
                            If (InStr(1, slInTime, "PM", vbTextCompare) = 0) And (InStr(1, slInTime, "AM", vbTextCompare) = 0) Then
                                ilPos = InStr(1, slInTime, "N", vbTextCompare)
                                If ilPos > 0 Then
                                    slInTime = Left(slInTime, ilPos - 1) & "PM"
                                End If
                                ilPos = InStr(1, slInTime, "M", vbTextCompare)
                                If ilPos > 0 Then
                                    slInTime = Left(slInTime, ilPos - 1) & "AM"
                                End If
                                ilPos = InStr(1, slInTime, "P", vbTextCompare)
                                If ilPos > 0 Then
                                    slInTime = Left(slInTime, ilPos - 1) & "PM"
                                End If
                                ilPos = InStr(1, slInTime, "A", vbTextCompare)
                                If ilPos > 0 Then
                                    slInTime = Left(slInTime, ilPos - 1) & "AM"
                                End If
                            End If
                            SQLQuery = SQLQuery & "astAirDate = '" & Format$(slInDate, sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery & "astAirTime = '" & Format$(slInTime, sgSQLTimeForm) & "', "
                            SQLQuery = SQLQuery & "astAffidavitSource = '" & Trim(slSource) & "'"
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportByASTCode"
                                mImportByAstCode = False
                                Exit Function
                            Else
                                tmAirSpotInfo(llLoop).iUpdateComplete = True
                            End If
                        ElseIf slSpotType = "N" Then
                            'update status as not aired
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            SQLQuery = SQLQuery + "astStatus = 4" & ", "  'Neither the spot nor the program aired.
                            SQLQuery = SQLQuery & "astAffidavitSource = '" & Trim(slSource) & "' "
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportByASTCode"
                                mImportByAstCode = False
                                Exit Function
                            Else
                                tmAirSpotInfo(llLoop).iUpdateComplete = True
                            End If
                            ilRet = mImportMissedReason(llLoop)
                        ElseIf slSpotType = "X" Then
                            'update status as not aired and no makegood
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            SQLQuery = SQLQuery + "astStatus = 14" & ", "  'Neither the spot nor the program aired.
                            SQLQuery = SQLQuery & "astAffidavitSource = '" & Trim(slSource) & "' "
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                GoSub ErrHand:
                            Else
                                tmAirSpotInfo(llLoop).iUpdateComplete = True
                            End If
                        ElseIf slSpotType = "P" Then
                            'update status as not aired
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            'SQLQuery = SQLQuery + "astStatus = 8"   'Program aired spot did not.
                            'D.S. 11/6/08
                            'Affiliate Meeting Decisions item 11) b-ii-d
                            SQLQuery = SQLQuery + "astStatus = 4" & ", "
                            SQLQuery = SQLQuery & "astAffidavitSource = '" & Trim(slSource) & "' "
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportByASTCode"
                                mImportByAstCode = False
                                Exit Function
                            Else
                                tmAirSpotInfo(llLoop).iUpdateComplete = True
                            End If
                            ilRet = mImportMissedReason(llLoop)
                        ElseIf slSpotType = "G" Then
                            'game was not carried
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            SQLQuery = SQLQuery + "astStatus = 8" & ", "   'Not carried or aired
                            SQLQuery = SQLQuery & "astAffidavitSource = '" & Trim(slSource) & "' "
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportByASTCode"
                                mImportByAstCode = False
                                Exit Function
                            Else
                                tmAirSpotInfo(llLoop).iUpdateComplete = True
                            End If
                        Else
                            'update status as not aired
                            SQLQuery = "UPDATE ast SET "
                            SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
                            SQLQuery = SQLQuery + "astStatus = 4" & ", "       'NA-Other
                            SQLQuery = SQLQuery & "astAffidavitSource = '" & Trim(slSource) & "' "
                            SQLQuery = SQLQuery + " WHERE (astCode = " & llAstCode & ")"
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportByASTCode"
                                mImportByAstCode = False
                                Exit Function
                            Else
                                tmAirSpotInfo(llLoop).iUpdateComplete = True
                            End If
                            ilRet = mImportMissedReason(llLoop)
                        End If
                        If llLoop <= UBound(tmAirSpotInfo) Then
                            'If Not myEnt.Add(tmAirSpotInfo(llLoop).sFeedDate, tmAstInfo(llLoop).lGsfCode, Ingested) Then
                            '    'myErrors.WriteWarning myEnt.ErrorMessage
                            'End If
                    Else
                            llLoop = llLoop
                        End If
                    Else
                        'gLogMsg "No match was found for this AST Code.", "WebImportLog.Txt", False
                        'gLogMsg "A match by feed date and time will be attempted.", "WebImportLog.Txt", False
                        imNextPassNeeded = True
                    End If
                Else
                    imNextPassNeeded = True
                End If
                ilRet = 0
            End If
        End If
    Next llLoop
                
'    If rstOpen Then
'        rst.Close
'    End If

    mImportByAstCode = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mImportByAstCode"
    mImportByAstCode = False
    Exit Function
End Function

Private Function mImportByFeedDate(sSunDate As String, sMonDate As String) As Integer

    'Created by D.S. June 2007
    'Try to find the AST by looking at the Feed Date and Feed Time.  If we find it then mark it as found in the array.
    
    Dim ilLoop As Integer
    'Dim rst As ADODB.Recordset
    Dim ilFound  As Integer
    Dim ilIndex As Integer
    Dim rstOpen As Integer
    Dim ilRet As Integer
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim ilAst As Integer
    
    On Error GoTo ErrHand
    
    mImportByFeedDate = False
    imImportByAstCodeAgain = False
    imNextPassNeeded = False
    
    rstOpen = False
    For ilLoop = LBound(tmAirSpotInfo) To UBound(tmAirSpotInfo) - 1 Step 1
        If tmAirSpotInfo(ilLoop).iFound = False Then
            'we need to add a sql call to att to get all of the att codes based on the vefCode and the ShttCode
            'that we already have:
            
            'tmAirSpotInfo(ilLoop).iShfCode
            'tmAirSpotInfo(ilLoop).iVefCode
            
            'loop through the result set using the tmAirSpotInfo(ilLoop).lAtfCode first.  If all spots are not posted
            'then look at the other agreements.  This way if the spots are on the web with a different attCode we can still
            'post them.  We would seach through the ast file, but there are no keys on the veh or the shtt.
        
            '12/13/13: Obtain Pledge information from Dat.  See change below
            'SQLQuery = "Select astCode, astShfCode, astAtfCode, astVefCode, astPledgeDate, astPledgeStatus, astFeedTime, astStatus FROM ast "
        
            'D.S. 11/4/15 take out below SQL to Do while gets replaced with a For loop on tmAstInfo and only process if iCpStatus = 0
            'for llAst = lBound tmAstInfo to ubound -1 step 1
            'if tmAstInfo(llAst).iCpstatus = 0 then

'            SQLQuery = "Select * FROM ast "
'            SQLQuery = SQLQuery & "Where astFeedDate = '" & Format$(tmAirSpotInfo(ilLoop).sFeedDate, sgSQLDateForm) & "' "
'            SQLQuery = SQLQuery & "And astCPStatus = 0 And astAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
'            Set rst = gSQLSelectCall(SQLQuery)
'            rstOpen = True
'
'            If Not rst.EOF Then
'                'We found the feed date now see if the time exists
'                Do While Not rst.EOF
                    
            For ilAst = 0 To UBound(tmAstInfo) - 1 Step 1
                DoEvents
                If tmAstInfo(ilAst).iCPStatus = 0 Then
                    '12/13/13: Obtain Pledge information from Dat
                    'tlDatPledgeInfo.lAttCode = rst!astAtfCode
                    tlDatPledgeInfo.lAttCode = tmAstInfo(ilAst).lAttCode
                    
                    'tlDatPledgeInfo.lDatCode = rst!astDatCode
                    tlDatPledgeInfo.lDatCode = tmAstInfo(ilAst).lDatCode
                    
                    'tlDatPledgeInfo.iVefCode = rst!astVefCode
                    tlDatPledgeInfo.iVefCode = tmAstInfo(ilAst).iVefCode
                    
                    'tlDatPledgeInfo.sFeedDate = Format(rst!astFeedDate, "m/d/yy")
                    tlDatPledgeInfo.sFeedDate = tmAstInfo(ilAst).sFeedDate
                    
                    'tlDatPledgeInfo.sFeedTime = Format(rst!astFeedTime, "hh:mm:ssam/pm")
                    tlDatPledgeInfo.sFeedTime = tmAstInfo(ilAst).sFeedTime
                    
                    ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
                    'If tgStatusTypes(gGetAirStatus(rst!astPledgeStatus)).iPledged <> 2 Then
                    If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo.iPledgeStatus)).iPledged <> 2 Then
                        'If Format(rst!astFeedTime, "hh:mm:ssam/pm") = tmAirSpotInfo(ilLoop).sFeedTime Then
                        
                        'If gTimeToLong(Format(rst!astFeedTime, "hh:mm:ssam/pm"), False) = gTimeToLong(tmAirSpotInfo(ilLoop).sFeedTime, False) Then
                        If gTimeToLong(Format(tmAstInfo(ilAst).sFeedTime, "hh:mm:ssam/pm"), False) = gTimeToLong(tmAirSpotInfo(ilLoop).sFeedTime, False) Then
                            ilFound = False
                            For ilIndex = LBound(tmAirSpotInfo) To UBound(tmAirSpotInfo) - 1 Step 1
                                If ilIndex <> ilLoop Then
                                    'If tmAirSpotInfo(ilIndex).lAstCode = rst!astCode Then
                                    If tmAirSpotInfo(ilIndex).lAstCode = tmAstInfo(ilAst).lCode Then
                                         ilFound = True
                                         Exit For
                                    End If
                                End If
                            Next ilIndex
                            
                            If Not ilFound Then
                                'tmAirSpotInfo(ilLoop).lAstCode = rst!astCode
                                tmAirSpotInfo(ilLoop).lAstCode = tmAstInfo(ilAst).lCode
                                tmAirSpotInfo(ilLoop).iFound = True
                                imImportByAstCodeAgain = True
                                lmResolvedBytime = lmResolvedBytime + 1
                                'Exit Do
                                tmAstInfo(ilAst).iCPStatus = 1
                                Exit For
                                    'D.S. 11/4/15 replace exit do with exit for above, also set tmAstInfo(llAst).iCpstatus = 1
                            End If
                        End If
                    End If
                        'D.S. 11/4/15 movenext and loop below gets replaced with next llAst
                    'rst.MoveNext
                'Loop
                End If
            Next ilAst
                Else
'                'We did not find the feed date with any agreement for that station/vehicle
                imNextPassNeeded = True
            End If
                If Not ilFound Then
                    imNextPassNeeded = True
        End If
    Next ilLoop
    If rstOpen Then
    rst.Close
    End If
    mImportByFeedDate = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mImportByFeedDate"
    Exit Function
End Function

Private Function mFillAirSpotInfo(llIdx As Long, iGetNewLine As Integer) As Integer

    'Created by D.S. June 2007
    
    'D.S. 6/11/12 pretty much re-written
    'Used Split function instead of gParseCDFields, added binary search for att info
    'Pump the tmAirSpotInfo array elements with the data in slFields

    'Import Record Fileds, the values for smFields(?)
    '1. attCode"
    '2. "Advt"
    '3. "Prod"
    '4. "PledgeStartDate1"
    '5. "PledgeEndDate"
    '6. "PledgeStartTime"
    '7. "PledgeEndTime"
    '8. "SpotLen"
    '9. "Cart"
    '10. "ISCI"
    '11. "CreativeTitle"
    '12. "astCode"
    '13. "ActualAirDate1"
    '14. "ActualAirTime1"
    '15. "statusCode"
    '16. "FeedDate"
    '17. "FeedTime"
    '18. Rec Type
    '19. Missed Reason
    '20. Original Ast Code
    '21. NewAstCode
    '22. gsfCode
    '23. Source - By whom was it posted. Vendor? Manually? Vendor then over written manually?

    Dim llLoop As Long
    Dim slLine As String
    Dim ilRet As Integer
    Dim ilTemp As Integer
    Dim llTemp As Long
    Dim slTemp As String
    Dim ilPos As Integer
    Dim slRecType As String
    Dim iNumRecCols As Integer
    Dim ilFillAirSpotInfo As Integer
    Dim slSource As String
    
    On Error GoTo ErrHand
    ilFillAirSpotInfo = 0
    mFillAirSpotInfo = 0
    If iGetNewLine Then
        Do
            If EOF(hmFrom) Then
                Exit Do
            End If
            Line Input #hmFrom, slLine
            If ilRet = 62 Then
                ilRet = 0
            End If
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Function
                End If
                slTemp = slLine
                slLine = Replace(slTemp, """, """, """,""")
                slLine = gStripCommaDG(slLine)
                'D.S. 07/26/12 On the Split below the """,""" means that it's actually looking for "," (all three chars) for its delimiter rather than just the comma.
                smRecordsArray = Split(slLine, """,""")
                'Sanity check on array rather than a string
                If Not IsArray(smRecordsArray) Then
                    mFillAirSpotInfo = 1
                    Exit Function
                End If
                'Sanity check on number of elements in the array
                'If Not (UBound(smRecordsArray) = 21 Or UBound(smRecordsArray) = 16) Then
                If Not (UBound(smRecordsArray) = 23 Or UBound(smRecordsArray) = 22 Or UBound(smRecordsArray) = 21 Or UBound(smRecordsArray) = 16) Then
                    Exit Function
                Else
                    iNumRecCols = UBound(smRecordsArray)
                End If
                'D.S. 10/26/15  I found that we were getting a mix of old and new record layouts from the web in a single import.  I was only checking the header to
                'determine old or new layouts.  I've remived that code and now check every line to verify.
                If UBound(smRecordsArray) = 16 Then
                    imOldImpLayout = True
                Else
                    imOldImpLayout = False
                End If
                lmTotalSpots = lmTotalSpots + 1
                ilFillAirSpotInfo = 1
                If lmTotalSpots Mod 1000 = 0 Then
                    gLogMsg "Processing Record " & lmTotalSpots, "WebImportLog.Txt", False
                    SetResults "Processing Record " & lmTotalSpots, 0
                End If
            Else
                If sgCommand = "/ReImport" Then
                    Exit Do
                End If
            End If
        Loop While Len(slLine) = 0
        If ilFillAirSpotInfo = 0 Then
            Exit Function
    End If
    End If
    
    tmAirSpotInfo(llIdx).lAtfCode = gGetDataNoQuotes(smRecordsArray(0))
        'D.S. 11/4/12 this is the src attcode
    If tmAirSpotInfo(llIdx).lAtfCode < 0 Then
        tmAirSpotInfo(llIdx).lAtfCode = gGetDataNoQuotes(smRecordsArray(21))
    End If
    llTemp = gBinarySearchAtt(tmAirSpotInfo(llIdx).lAtfCode)
    If llTemp <> -1 Then
        tmAirSpotInfo(llIdx).iShfCode = tgAttInfo1(llTemp).attShttCode
        tmAirSpotInfo(llIdx).iVefCode = tgAttInfo1(llTemp).attvefCode
    Else
        tmAirSpotInfo(llIdx).iShfCode = gGetShttCodeFromAttCode(CStr(tmAirSpotInfo(llIdx).lAtfCode))
        tmAirSpotInfo(llIdx).iVefCode = gGetVehCodeFromAttCode(CStr(tmAirSpotInfo(llIdx).lAtfCode))
    End If
    tmAirSpotInfo(llIdx).sAdvt = gGetDataNoQuotes(smRecordsArray(1))
    tmAirSpotInfo(llIdx).sProd = gGetDataNoQuotes(smRecordsArray(2))
    tmAirSpotInfo(llIdx).sPledgeStartDate1 = gGetDataNoQuotes(smRecordsArray(3))
    tmAirSpotInfo(llIdx).sPledgeEndDate = gGetDataNoQuotes(smRecordsArray(4))
    tmAirSpotInfo(llIdx).sPledgeStartTime = gGetDataNoQuotes(smRecordsArray(5))
    tmAirSpotInfo(llIdx).sPledgeEndTime = gGetDataNoQuotes(smRecordsArray(6))
    tmAirSpotInfo(llIdx).iSpotLen = gGetDataNoQuotes(smRecordsArray(7))
    tmAirSpotInfo(llIdx).sCart = gGetDataNoQuotes(smRecordsArray(8))
    tmAirSpotInfo(llIdx).sISCI = gGetDataNoQuotes(smRecordsArray(9))
    tmAirSpotInfo(llIdx).sCreativeTitle = gGetDataNoQuotes(smRecordsArray(10))
    tmAirSpotInfo(llIdx).lAstCode = gGetDataNoQuotes(smRecordsArray(11))
    tmAirSpotInfo(llIdx).sActualAirDate1 = Format(gGetDataNoQuotes(smRecordsArray(12)), sgSQLDateForm)
    tmAirSpotInfo(llIdx).sActualAirTime1 = Format(gGetDataNoQuotes(smRecordsArray(13)), "hh:mm:ssam/pm")
    tmAirSpotInfo(llIdx).sStatusCode = gGetDataNoQuotes(smRecordsArray(14))
    If smRecordsArray(15) = "" Then
        tmAirSpotInfo(llIdx).sFeedDate = Format(gGetDataNoQuotes(smRecordsArray(3)), sgSQLDateForm)
        tmAirSpotInfo(llIdx).sFeedTime = Format(gGetDataNoQuotes(smRecordsArray(5)), "hh:mm:ssam/pm")
    Else
        tmAirSpotInfo(llIdx).sFeedDate = Format(gGetDataNoQuotes(smRecordsArray(15)), sgSQLDateForm)
        tmAirSpotInfo(llIdx).sFeedTime = Format(gGetDataNoQuotes(smRecordsArray(16)), "hh:mm:ssam/pm")
    End If
    
    tmAirSpotInfo(llIdx).bIsciChngFlag = False
    If Not imOldImpLayout Then
        'If the len of tmAirSpotInfo(llLoop).sRecType) > 1 then we have a G rectype.  Rectype G is a record that the ISCI code
        'has been changed from what was originally exported.
'        If Len(Trim(Trim(smRecordsArray(17)))) > 1 Then
'            ilPos = InStr(1, smRecordsArray(17), "G", vbTextCompare)
'            If ilPos = 2 Then
'                slRecType = Trim$(Left$(smRecordsArray(17), 1))
'            Else
'                slRecType = Trim$(right$(smRecordsArray(17), 1))
'            End If
'            If slRecType <> "M" And slRecType <> "R" And slRecType <> "B" Then
'                tmAirSpotInfo(llIdx).bIsciChngFlag = True
'            End If
'        Else
'            slRecType = smRecordsArray(17)
'        End If
        
        
        If Len(Trim(Trim(smRecordsArray(17)))) > 1 Then
            ilPos = InStr(1, smRecordsArray(17), "G", vbTextCompare)
            If ilPos = 2 Then
                slRecType = Trim$(Left$(smRecordsArray(17), 1))
            Else
                slRecType = Trim$(right$(smRecordsArray(17), 1))
            End If
            If slRecType <> "M" And slRecType <> "R" And slRecType <> "B" Then
                tmAirSpotInfo(llIdx).bIsciChngFlag = True
            End If
        Else
            If Trim(smRecordsArray(17)) = "G" Then
                tmAirSpotInfo(llIdx).bIsciChngFlag = True
            End If
            slRecType = smRecordsArray(17)
        End If
        'tmAirSpotInfo(llIdx).sRecType = gGetDataNoQuotes(smRecordsArray(17))   'M = make good, B = bonus, R = replacement
        tmAirSpotInfo(llIdx).sRecType = slRecType
        tmAirSpotInfo(llIdx).iMissedReason = gGetDataNoQuotes(smRecordsArray(18))
        tmAirSpotInfo(llIdx).lOrgAstCode = gGetDataNoQuotes(smRecordsArray(19))
        tmAirSpotInfo(llIdx).lNewAstCode = gGetDataNoQuotes(smRecordsArray(20))
        If iNumRecCols > 21 Then
            tmAirSpotInfo(llIdx).lgsfCode = gGetDataNoQuotes(smRecordsArray(22))
        Else
            tmAirSpotInfo(llIdx).lgsfCode = -1
        End If
        If iNumRecCols > 22 Then
            slSource = gGetDataNoQuotes(smRecordsArray(23))
            If slSource = "" Then
                slSource = "UD"  ' UD = Undefined
            End If
            tmAirSpotInfo(llIdx).sVendorSource = slSource
        End If
    End If
    tmAirSpotInfo(llIdx).iFound = False
    tmAirSpotInfo(llIdx).iUpdateComplete = False
    mFillAirSpotInfo = 1
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mFillAirSpotInfo"
    Exit Function
End Function

Private Function mUpdateCptt() As Integer

    'Created by D.S. June 2007
    'Set the CPTT week's value
    
    Dim ilLoop As Integer
    Dim llPrevAttCode As Long
    Dim slSDate As String
    Dim slMoDate As String
    Dim slSuDate As String
    Dim ilStatus As Integer
    Dim llVeh As Long
    Dim llCount As Long
    
    On Error GoTo ErrHand
    
    mUpdateCptt = False
    'Set any Not Aired to received as they are not exported
    llPrevAttCode = -1
    
    'Typically we go one record too far when looping through the import web spots file we recieved from the web.
    'But, when we hit the end of the file we have NOT gone one record too far.  So, if we are NOT at the EOF we
    'are good on the record count.  If we are at the EOF then we need to subtract one from the count because we
    'haven't gone one too far.
    '8/3/11: Ignore imEOF
    'If imEOF Then
        llCount = UBound(tmAirSpotInfo) - 1
    'Else
    '    llCount = UBound(tmAirSpotInfo)
    'End If
    
    'For ilLoop = 0 To UBound(tmAirSpotInfo) Step 1
    For ilLoop = 0 To llCount Step 1
        If llPrevAttCode <> tmAirSpotInfo(ilLoop).lAtfCode Then
            DoEvents
            'slSDate = tmAirSpotInfo(ilLoop).sStartDate
            slSDate = tmAirSpotInfo(ilLoop).sFeedDate
            slMoDate = gAdjYear(gObtainPrevMonday(slSDate))
            Do
                slSuDate = DateAdd("d", 6, slMoDate)
                For ilStatus = 0 To UBound(tgStatusTypes) Step 1
                    If (tgStatusTypes(ilStatus).iPledged = 2) Then
                        SQLQuery = "UPDATE ast SET "
                        If imNoSpotsAired Then
                            'D.S. 10/26/19
                            SQLQuery = SQLQuery + "astLen = " & tmAirSpotInfo(ilLoop).iSpotLen & ","
                            SQLQuery = SQLQuery + "astCPStatus = " & "2"    'No spots aired
                        Else
                            SQLQuery = SQLQuery + "astLen = " & tmAirSpotInfo(ilLoop).iSpotLen & ","
                            SQLQuery = SQLQuery + "astCPStatus = " & "1"    'Received
                        End If
                        SQLQuery = SQLQuery + " WHERE (astAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                        'Check for any spots that have not aired - astCPStatus = 0 = not aired
                        SQLQuery = SQLQuery + " AND astCPStatus = 0"
                        SQLQuery = SQLQuery + " AND Mod( astStatus, " & 100 & ") = " & tgStatusTypes(ilStatus).iStatus
                        SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
                        'D.S. 08/29/11 Running out of transactions and causing errors
                        'cnn.BeginTrans
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/13/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "WebImportLog.txt", "WebImportSchdSpot-mUpdateCPTT"
                            mUpdateCptt = False
                            Exit Function
                        End If
                        'cnn.CommitTrans
                    End If
                Next ilStatus
                slMoDate = DateAdd("d", 7, slMoDate)
            'Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sEndDate))
            Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sFeedDate))
        End If
        llPrevAttCode = tmAirSpotInfo(ilLoop).lAtfCode
    Next ilLoop
    'Determine if CPTTStatus should to set to 0=Partial or 1=Completed
    llPrevAttCode = -1
    'For ilLoop = 0 To UBound(tmAirSpotInfo) Step 1
    For ilLoop = 0 To llCount Step 1
        If llPrevAttCode <> tmAirSpotInfo(ilLoop).lAtfCode Then
            DoEvents
            'slSDate = tmAirSpotInfo(ilLoop).sStartDate
            slSDate = tmAirSpotInfo(ilLoop).sFeedDate
            slMoDate = gAdjYear(gObtainPrevMonday(slSDate))
            Do
                slSuDate = DateAdd("d", 6, slMoDate)
                SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
                SQLQuery = SQLQuery + " AND astAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                SQLQuery = SQLQuery + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                Set rst = gSQLSelectCall(SQLQuery)
                If rst.EOF Then
                    'Set CPTT as complete
                    SQLQuery = "UPDATE cptt SET "
                    llVeh = gBinarySearchVef(CLng(imVefCode))
                    If llVeh <> -1 Then
                        If (tgVehicleInfo(llVeh).sVehType = "G") And (DateValue(slSuDate) > DateValue(Format$(gNow(), "DDDDD"))) Then
                            SQLQuery = SQLQuery + "cpttStatus = 0" & ", " 'Partial
                            SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery + "cpttPostingStatus = 1" 'Partial
                        Else
                            If imNoSpotsAired Then
                                If mCheckAnyAired(tmAirSpotInfo(ilLoop).lAtfCode, slMoDate, slSuDate) Then
                                    imNoSpotsAired = False
                                End If
                            End If
                            If imNoSpotsAired Then
                                SQLQuery = SQLQuery + "cpttStatus = 2" & ", " 'Complete
                            Else
                                SQLQuery = SQLQuery + "cpttStatus = 1" & ", " 'Complete
                            End If
                            SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                            SQLQuery = SQLQuery + "cpttPostingStatus = 2"  'Complete
                        End If
                    Else
                        If imNoSpotsAired Then
                            If mCheckAnyAired(tmAirSpotInfo(ilLoop).lAtfCode, slMoDate, slSuDate) Then
                                imNoSpotsAired = False
                            End If
                        End If
                        If imNoSpotsAired Then
                            SQLQuery = SQLQuery + "cpttStatus = 2" & ", " 'Complete
                        Else
                        SQLQuery = SQLQuery + "cpttStatus = 1" & ", " 'Complete
                        End If
                        SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                        SQLQuery = SQLQuery + "cpttPostingStatus = 2"  'Complete
                    End If
                    SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                    SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "WebImportLog.txt", "WebImportSchdSpot-mUpdateCPTT"
                        mUpdateCptt = False
                        Exit Function
                    End If
                Else
                    'Set CPTT as partial
                    SQLQuery = "UPDATE cptt SET "
                    SQLQuery = SQLQuery + "cpttStatus = 0" & ", " 'Partial
                    SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    SQLQuery = SQLQuery + "cpttPostingStatus = 1" 'Partial
                    SQLQuery = SQLQuery + " WHERE cpttAtfCode = " & tmAirSpotInfo(ilLoop).lAtfCode
                    SQLQuery = SQLQuery + " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "WebImportLog.txt", "WebImportSchdSpot-mUpdateCPTT"
                        mUpdateCptt = False
                        Exit Function
                    End If
                End If
                slMoDate = DateAdd("d", 7, slMoDate)
            'Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sEndDate))
            Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(tmAirSpotInfo(ilLoop).sFeedDate))
        End If
        llPrevAttCode = tmAirSpotInfo(ilLoop).lAtfCode
    Next ilLoop
    gFileChgdUpdate "cptt.mkd", True
    mUpdateCptt = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mUpdateCptt"
    Exit Function
End Function

Private Function mProcessWebWorkStatusResults(sFileName As String, sIniValue As String) As Boolean

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
    
    On Error GoTo FileErrHand:
    'hlFrom = FreeFile
    'ilRet = 0
    'Open slLocation For Input Access Read As hlFrom
    ilRet = gFileOpen(slLocation, "Input Access Read", hlFrom)
    If ilRet <> 0 Then
        gMsgBox "Error: frmWebExportSchdSpot-mProcessWebWorkStatusResults was unable to open the file."
        GoTo ErrHand
    End If
    
    'Skip past the header record
    Input #hlFrom, smFileName, smStatus, smMsg1, smMsg2, smDTStamp
    Input #hlFrom, smFileName, smStatus, smMsg1, smMsg2, smDTStamp
    
    slTemp = smMsg2
    
    ilLen = Len(slTemp)
    ilPos = InStr(slTemp, ":")
    llCount = Val(Mid$(slTemp, ilPos + 1, ilLen))
'    If ilPos > 0 Then
'        If InStr(slTemp, "Total Comments Imported:") Then
'            lmWebTtlComments = lmWebTtlComments + llCount
'        End If
'
'        If InStr(slTemp, "Total Headers Imported:") Then
'            lmWebTtlHeaders = lmWebTtlHeaders + llCount
'        End If
'
'        If InStr(slTemp, "Total Records Processed:") Then
'            lgWebTtlSpots = lgWebTtlSpots + llCount
'        End If
'
'        If InStr(slTemp, "Total Emails Sent:") Then
'            lgWebTtlEmail = lgWebTtlEmail + llCount
'        End If
'    End If
    
    'Cover the case that the Web Server times out and does not create the second line in the file
    If smStatus = "Status" Then
        smStatus = "1"
        gLogMsg "Warning: " & "Had to Set smStatus to 1 because the Work Status File Only had the Header in it.", "WebImportLog.Txt", False
    End If
    
    Close hlFrom
    mProcessWebWorkStatusResults = True
    Exit Function

FileErrHand:
    Close hlFrom
    mProcessWebWorkStatusResults = True
    'Cover the case that the Web Server times out and does not create the second line in the file
    If smStatus = "Status" Then
        smStatus = "1"
        gLogMsg "Warning: FileErrHand " & "Had to Set smStatus to 1 because the Work Status File Only had the Header in it.", "WebImportLog.Txt", False
    End If
    
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mProcessWebWorkStatusResults"
    Exit Function
End Function

'***************************************************************************************
' JD 08-22-2007
' This function was added to handle a special case occurring in the function
' mCheckWebWorkStatus. We believe a network error is causing the error handler
' to fire. Adding retry code to the function mCheckWebWorkStatus itself did not
' seem feasable because we did not know where the error was actually occuring and
' simplying calling a resume next could cause even more trouble.
'
'***************************************************************************************
Private Function mExCheckWebWorkStatus(sFileName As String) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilLine As Integer
    Dim ilErrNo As Integer
    Dim slDesc As String

    On Error GoTo Err_Handler
    mExCheckWebWorkStatus = -1
    For ilLoop = 1 To 10
        ilRet = mCheckWebWorkStatus(sFileName)
        mExCheckWebWorkStatus = ilRet
        If ilRet <> -2 Then ' Retry only when this status is returned.
            Exit Function
        End If
        gLogMsg "mExCheckWebWorkStatus is retrying due to an error in mCheckWebWorkStatus", "WebImpRetryLog.txt", False
        DoEvents
        Sleep 2000  ' Delay for two seconds when retrying.
    Next
    If ilRet = -2 Then
        ilRet = -1  ' Keep the original error of -1 so all callers can process the error normally.
        gMsg = "A timeout has occured in frmWebImportAiredSpot - mExCheckWebWorkStatus"
        gLogMsg gMsg, "WebImpRetryLog.txt", False
        gLogMsg " ", "WebImpRetryLog.txt", False
    End If
    Exit Function

Err_Handler:
    Screen.MousePointer = vbDefault
    mExCheckWebWorkStatus = -1
    gMsg = ""
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gMsg = "Error: " & "A general error has occured in frmWebImportAiredSpot - mExCheckWebWorkStatus: " & "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc
    gLogMsg gMsg, "WebImportLog.txt", False
    gLogMsg " ", "WebImportLog.txt", False
    Exit Function
End Function

Private Function mCheckWebWorkStatus(sFileName As String) As Integer
 
    'D.S. 6/22/05
    
    'input - sFilemane is the unique file name that is the key into the web
    'server database to check it's status
    
    'Web Server Status - 0 = Done, 1 = Working and 2 = Error
    
    'Loop while the web server is busy processing spots and emails
    'Check the server every 10 seconds Report status
    
    Dim sFTPAddress As String
    'Dim ilRet As Integer
    Dim llWaitTime As Long
    Dim ilModResult As Integer
    Dim imStatus As Integer
    Dim slResult As String
    Dim llNumRows As Long
    Dim ilTimedOut As Integer
    Dim slTemp As String
    
    'Debug information
    Dim ilLine As Integer
    Dim slDesc As String
    Dim ilErrNo As Integer
    
    'Number of Seconds to Sleep
    Const clNumSecsToSleep As Long = 10
    
    Const clSleepValue As Long = clNumSecsToSleep * cmOneSecond
    
    'Assuming clNumSecsToSleep is 10 then a mod value of 6 would
    'be 6 loops at 10 seconds each or 1 minute
    Const clModValue As Integer = 6
    
    On Error GoTo ErrHand
    
    mCheckWebWorkStatus = False
    If Not gHasWebAccess() Then
        Exit Function
    End If
    
    Call gLoadOption(sgWebServerSection, "FTPAddress", sFTPAddress)
    llWaitTime = 0
    imStatus = 1
    Do While imStatus = 1 And llWaitTime < 1350 'We will wait 45 minutes based on 1350 - 1350/60 * 2 seconds
        DoEvents
        Sleep clSleepValue
        SQLQuery = "Select Count(*) from WorkStatus Where FileName = " & "'" & sFileName & "'"
        llNumRows = gExecWebSQLWithRowsEffected(SQLQuery)
        
        If llNumRows = -1 Then
            'An error was returned
            imStatus = 2
        End If
        If llNumRows > 0 Then
            SQLQuery = "Select FileName, Status, Msg1, Msg2, DTStamp from WorkStatus Where FileName = " & "'" & sFileName & "'"
            'Get the status information from the web server database and write it to a file
            Call gRemoteExecSql(SQLQuery, smWebWorkStatus, "WebImports", True, True, 30)
            DoEvents
            Call mProcessWebWorkStatusResults(smWebWorkStatus, "WebImports")
            llWaitTime = llWaitTime + 1
            ilModResult = llWaitTime Mod clModValue
            imStatus = CInt(smStatus)
            'Handle Web Error Condition
            If imStatus = 2 Then
                If StrComp(Trim$(smMsg1), "WARN: An export is already running. Commit needed. Operation aborted.") = 0 Then
                    'SetResults "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), 0
                    mCheckWebWorkStatus = False
                    imTerminate = True
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
                
                slTemp = Left$(Trim$(smMsg1), 35)
                If StrComp(Trim$(slTemp), "ERROR: Export Count does not match.") = 0 Then
                    SetResults "   " & smMsg1, 0
                    mCheckWebWorkStatus = False
                    imTerminate = True
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            
                gLogMsg "Error: " & "The Web Server Returned an ERROR. See Below. ", "WebImportLog.Txt", False
                gLogMsg "   " & "Error: " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebImportLog.Txt", False
                Call gEndWebSession("WebImportLog.Txt")
                mCheckWebWorkStatus = False
                Exit Function
            End If
            If ilModResult = 0 And imStatus = 1 Then
                DoEvents
                SetResults "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), 0
                'SetResults "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), 0
                gLogMsg smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebImportLog.Txt", False
                gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebImportLog.Txt", False
                DoEvents
            End If
        End If
    Loop
    
    If llWaitTime >= 1350 Then
        'We timed out
        gLogMsg "Error: " & "   " & "A timeout occured while waiting on the web server for a response.", "WebImportLog.Txt", False
        SetResults "A timeout occured waiting on a web server response.", RGB(255, 0, 0)
        Call gEndWebSession("WebImportLog.Txt")
        mCheckWebWorkStatus = False
        Exit Function
        
    End If
    
    'Show the final message with the totals of spots imported an emails sent
    'Call mProcessWebWorkStatusResults(smWebWorkStatus, "WebExports")

    imStatus = CInt(smStatus)
    'Handle Web Error Condition
    If imStatus = 2 Then
        gLogMsg "Error: " & "   " & "The Web Server Returned an ERROR. See Below. ", "WebImportLog.Txt", False
        gLogMsg "Error: " & "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebImportLog.Txt", False
        Call gEndWebSession("WebImportLog.Txt")
        mCheckWebWorkStatus = False
        Exit Function
    End If
    SetResults "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), 0
    'SetResults "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), 0
    gLogMsg "   " & smMsg1 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebImportLog.Txt", False
    gLogMsg "   " & smMsg2 & " - " & Format(smDTStamp, "h:mm:ss am/pm"), "WebImportLog.Txt", False
    mCheckWebWorkStatus = True
    
Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    mCheckWebWorkStatus = -2
    gMsg = ""
    ilLine = Erl
    ilErrNo = Err.Number
    slDesc = Err.Description
    gMsg = "A general error has occured in frmWebImportAiredSpot - mCheckWebWorkStatus: " & "  ilLine = " & ilLine & " ilErrNo = " & ilErrNo & " slDesc = " & slDesc
    gLogMsg "Error: " & gMsg, "WebImportLog.txt", False
    Exit Function
End Function

'
' Check to make sure the first record contains the field definitions.
' Only the first field name is looked at.
Private Function mCheckFile()
    On Error GoTo mImportSpotsErr:
    Dim slPathFileName As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim hlFH As Integer

    On Error GoTo mImportSpotsErr:
    
    mCheckFile = True
    Exit Function  ' NOT DONE. This needs to be done for all files.
    
    ' Need to check all the files.
    If Not OpenTextFile(hlFH, smWebSpots) Then
        mCheckFile = False
        Exit Function
    End If

    ' Read the column definition fields.
    Dim sATTCode, Advertiser, Product, PledgeStartDate, PledgeEndDate, PledgeStartTime, PledgeEndTime, SpotLen, Cart, ISCI, Title, astCode As String
    Input #hlFH, sATTCode, Advertiser, Product, PledgeStartDate, PledgeEndDate, PledgeStartTime, PledgeEndTime, SpotLen, Cart, ISCI, Title, astCode
    Close hlFH
    If Len(sATTCode) < 1 Or sATTCode <> "attCode" Then
        mCheckFile = False
    End If
    Exit Function

mImportSpotsErr:
    ilRet = Err.Number
    Resume Next

End Function

Private Sub SetResults(Msg As String, FGC As Long)
    lbcMsg.AddItem Msg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    lbcMsg.ForeColor = FGC
    DoEvents
End Sub

Private Sub mBuildFileNames()

    Dim slMsgFileName As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slTemp As String
    Dim slTemp2 As String
    
    ilRet = 0
        
    slTemp = gGetComputerName()
    If slTemp = "N/A" Then
        slTemp = "Unknown"
    End If
    
    smWebWorkStatus = "WebWorkStatus_" & slTemp & "_" & sgUserName & ".txt"
    slTemp = slTemp & "_" & sgUserName & "_" & Format(Now(), "yymmdd") & "_" & Format(Now(), "hhmmss") & ".txt"
    
    'debug
    'sgCommand = "/m"
    If (StrComp(sgCommand, "/m", vbTextCompare) = 0) Then
        smWebSpots = "WebSpots_" & "Manual.txt"
        smWebHeaders = "WebHeaders_" & "Manual.txt"
        smWebLogs = "WebLogs_" & "Manual.txt"
    Else
        smWebSpots = "WebSpots_" & slTemp
        smWebHeaders = "WebHeaders_" & slTemp
        smWebLogs = "WebLogs_" & slTemp
    End If

End Sub

Private Sub mWaitForWebLock()
    On Error GoTo ErrHandler
    Dim ilLoop As Integer
    Dim ilTotalMinutes As Integer
    Dim ilNotSaidWebServerWasBusy As Boolean
    Dim slLastMessage As String
    Dim slThisMessage As String
    Dim ilRow As Integer
    
    ilNotSaidWebServerWasBusy = False
    slLastMessage = "Nothing"
    While 1
        ilTotalMinutes = gStartWebSession("WebImportLog.Txt")
        If ilTotalMinutes = 0 Then
            'Start the Export Process
            'gLogMsg "Web Session Started Successfully", "WebImportLog.Txt", False
            Exit Sub
        End If
        If Not ilNotSaidWebServerWasBusy Then
            ilNotSaidWebServerWasBusy = True
            SetResults "The Server is Busy. Standby...", 0
            gLogMsg "The Server is Busy. Standby...", "WebImportLog.Txt", False
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
            lbcMsg.AddItem slThisMessage
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
    gMsg = "A general error has occured in frmWebImportAiredSpot-mWaitForWebLock: "
    gLogMsg "Error: " & gMsg & Err.Description & " Error #" & Err.Number & "; Line #" & Erl, "WebImportLog.Txt", False
End Sub

Private Sub DoEventsTimer_Timer()
    DoEvents
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
    If sgCommand = "/ReImport" Then
        Me.Left = -2 * Screen.Width
    End If
End Sub

Private Sub Form_Load()

    Dim iLoop As Integer
    Dim iZone As Integer
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    frmWebImportAiredSpot.Caption = "Web Aired Station Spots - " & sgClientName
    imAllClick = False
    imTerminate = False
    imImporting = False
    '10000
    lbcWebType.FontSize = 6
    If igDemoMode Then
        lbcWebType.Caption = "Demo Mode"
    ElseIf gIsTestWebServer() Then
        lbcWebType.Caption = "Test Website"
    End If
    Call gLoadOption(sgWebServerSection, "WebImports", smWebImports)
    smWebImports = gSetPathEndSlash(smWebImports, True)
    Screen.MousePointer = vbDefault
    DoEventsTimer.Interval = 500
    DoEventsTimer.Enabled = True
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    '8862
    tmAutoUpdateVendors = gAutoDeliveryVendors()
    If igAutoImport Then
        sgTimeZone = Left$(gGetLocalTZName(), 1)
'        D.S. 2/19/20 moved to cmdImport
'        tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
'        tmcSetTime.Enabled = True
'        gUpdateTaskMonitor 1, "ASI"
        ilRet = mGetSports()
        ilRet = gPopAll
        frmWebImportAiredSpot.Show
        cmdImport_Click
        gUpdateTaskMonitor 2, "ASI"
        cmdCancel_Click
    ElseIf sgCommand = "/ReImport" Then
        mReImport
        Exit Sub
    End If
    ilRet = gPopVff()
    'Dan 8/27/18 set myErrors! Still have issues...remove myErrors
'    Set myErrors = New CLogger
'    myErrors.LogPath = myErrors.CreateLogName(sgMsgDirectory & cmPathForgLogMsg)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmAirSpotInfo
    Erase tmTempAirSpotInfo
    Erase smRecordsArray
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    tmcSetTime.Enabled = False
    DoEventsTimer.Enabled = False
    Close #hmFrom
    Erase tmCPDat
    Erase tmAstInfo
    cprst.Close
    lat_recs.Close
    rst_Cpf.Close
    lst_rst.Close
   'Set myErrors = Nothing
    Set myEnt = Nothing
    Set frmWebImportAiredSpot = Nothing
    

End Sub


Public Function OpenTextFile(FH As Integer, PathFileName As String) As Boolean
    'On Error GoTo OpenTextFile:
    Dim ilRet As Integer
    
    OpenTextFile = True
    'FH = FreeFile
    'Open PathFileName For Input Access Read As FH
    ilRet = gFileOpen(PathFileName, "Input Access Read", FH)
    If ilRet <> 0 Then
        OpenTextFile = False
    End If
    Exit Function
    
'OpenTextFile:
'    ilRet = Err.Number
'    Resume Next
    
End Function

Private Function mEraseWebSpots(sCutoffDate As String) As Long
   
   'Jeff D. 08/20/08
   'Doug.S. Mods 08/22/08
   'Doug.S. Mods 02/3/09
    
   Dim llRowsEffected As Long
 
   On Error GoTo ErrHand
   ' Determine whether to use FeedDate or PledgeStartDate
   llRowsEffected = gExecWebSQLWithRowsEffected("Select Count(*) From Spots Where FeedDate is Null")
   If llRowsEffected > 0 Then
      llRowsEffected = gExecWebSQLWithRowsEffected("Select Count(*) From Spots Where PledgeStartDate <= '" & sCutoffDate & "'")
      If llRowsEffected = 0 Then
         ' No spots were found so exit now.
         mEraseWebSpots = 0
         Exit Function
      End If
      If llRowsEffected = -1 Then
         ' Error the function did not execute correctly.
         gLogMsg "Error: " & "frmWebImportAiredSpot-mEraseWebSpots: Select * From Spots Where PledgeStartDate <= '" & sCutoffDate & "'", "WebImportLog.Txt", False
         mEraseWebSpots = -1
         Exit Function
      End If
      
      llRowsEffected = gExecWebSQLWithRowsEffected("Delete From Spots Where PledgeStartDate <= '" & sCutoffDate & "'")
      If llRowsEffected = -1 Then
         ' Error the function did not execute correctly.
         mEraseWebSpots = -1
         gLogMsg "Error: " & "frmWebImportAiredSpot-mEraseWebSpots: Delete From Spots Where PledgeStartDate <= '" & sCutoffDate & "'", "WebImportLog.Txt", False
         Exit Function
      End If
      llRowsEffected = gExecWebSQLWithRowsEffected("Delete From SpotRevisions Where PledgeStartDate <= '" & sCutoffDate & "'")
      If llRowsEffected = -1 Then
         ' Error the function did not execute correctly.
         mEraseWebSpots = -1
         gLogMsg "Error: " & "frmWebImportAiredSpot-mEraseWebSpots: Delete From SpotRevisions Where PledgeStartDate <= '" & sCutoffDate & "'", "WebImportLog.Txt", False
         Exit Function
      End If
   Else
      llRowsEffected = gExecWebSQLWithRowsEffected("Select Count(*) From Spots Where FeedDate <= '" & sCutoffDate & "'")
      If llRowsEffected = 0 Then
         ' No spots were found so exit now.
         mEraseWebSpots = 0
         Exit Function
      End If
      If llRowsEffected = -1 Then
         ' Error the function did not execute correctly.
         mEraseWebSpots = -1
         gLogMsg "Error: " & "frmWebImportAiredSpot-mEraseWebSpots: Select Count(*) From Spots Where FeedDate <= '" & sCutoffDate & "'", "WebImportLog.Txt", False
         Exit Function
      End If
      
      llRowsEffected = gExecWebSQLWithRowsEffected("Delete From Spots Where FeedDate <= '" & sCutoffDate & "'")
      If llRowsEffected = -1 Then
         ' Error the function did not execute correctly.
         mEraseWebSpots = -1
         gLogMsg "Error: " & "frmWebImportAiredSpot-mEraseWebSpots: Delete From Spots Where FeedDate <= '" & sCutoffDate & "'", "WebImportLog.Txt", False
         Exit Function
      End If
      llRowsEffected = gExecWebSQLWithRowsEffected("Delete From SpotRevisions Where FeedDate <= '" & sCutoffDate & "'")
      If llRowsEffected = -1 Then
         ' Error the function did not execute correctly.
         mEraseWebSpots = -1
         gLogMsg "Error: " & "frmWebImportAiredSpot-mEraseWebSpots: Delete From SpotRevisions Where FeedDate <= '" & sCutoffDate & "'", "WebImportLog.Txt", False
         Exit Function
      End If
   End If
   mEraseWebSpots = llRowsEffected
  
   Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    SetResults "Error Erasing", RGB(255, 0, 0)
    SetResults "See Error Log", RGB(255, 0, 0)
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebImportAiredSpot-mEraseWebSpots: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "WebImportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    Exit Function
End Function

Private Function mTestToErase() As Long

    'D.S. 08/20/08
    'Test to see if the saved number of months to Erase is valid
    'Test to see if the saved date to erase up to is valid
    'If both are valid then call the web to erase
    'The web will try to Erase on each try, even if it's been done before
    'The web will report the number of rows affected
    
    Dim llRet As Long
    Dim rst_Temp As ADODB.Recordset
    Dim ilMonthsToRetainSpots As Integer
    Dim ilOkToErase As Integer
    Dim smInvEndStdMnth As String
    Dim slCuttOffDate As String
    
    On Error GoTo ErrHand
    
    mTestToErase = 0
    'Erase all of the WEB spots older than a given date
    SQLQuery = "Select spfRetainAffSpot, spfBLastStdMnth From SPF_Site_Options"
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    ilOkToErase = False
    If Not rst_Temp.EOF Then
        smInvEndStdMnth = Format(rst_Temp!spfBLastStdMnth, "yyyy-mm-dd")
        ilMonthsToRetainSpots = rst_Temp!spfRetainAffSpot
        'Sanity Check
        If smInvEndStdMnth <> "" Then
            If ilMonthsToRetainSpots = 0 Then
                'Set to default
                ilMonthsToRetainSpots = 24
            End If
            If ilMonthsToRetainSpots > 0 And ilMonthsToRetainSpots < 500 Then
                ilOkToErase = True
            Else
                gLogMsg "Warning: Number of months to retain spots was not valid. Number tested was: " & ilMonthsToRetainSpots, "WebImportLog.Txt", False
                gLogMsg "No Erasing was Attempted", "WebImportLog.Txt", False
                SetResults "*** Warning ***, Number of months to retain spots is set to: " & ilMonthsToRetainSpots, 0
                SetResults "    Please Correct in Traffic - Site. Number of Months Must be Greater than Zero", RGB(255, 0, 0)
            End If
        End If
        
        If ilOkToErase Then
            ilOkToErase = False
            SQLQuery = "Select safEarliestAffSpot, safLastArchRunDate From SAF_Schd_Attributes WHERE safVefCode = 0"
            Set rst_Temp = gSQLSelectCall(SQLQuery)
            If Not rst_Temp.EOF Then
                slCuttOffDate = Format(gObtainEndStd(DateAdd("d", (-31 * ilMonthsToRetainSpots) - 7, smInvEndStdMnth)), "yyyy-mm-dd")
                smEarliestSpottDate = Format$(rst_Temp!safEarliestAffSpot, "yyyy-mm-dd")
                'Sanity Check
                If gDateValue(smEarliestSpottDate) <= gDateValue(slCuttOffDate) Then
                    ilOkToErase = True
                End If
            End If
        End If
        rst_Temp.Close
    
        If ilOkToErase And smEarliestSpottDate <> "" Then
            SetResults "Checking for Spots Prior to: " & smEarliestSpottDate, 0
            gLogMsg "Checking for Spots Prior to: " & smEarliestSpottDate, "WebImportLog.Txt", False
            llRet = mEraseWebSpots(smEarliestSpottDate)
        End If
    End If
    mTestToErase = llRet
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mTestToErase"
    SetResults "Error Erasing", RGB(255, 0, 0)
    SetResults "See Error Log", RGB(255, 0, 0)
    Exit Function
End Function

Private Function mInitFTP() As Boolean

    Dim slTemp As String
    Dim ilRet As Integer
    Dim slSection As String
    
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
    Call gLoadOption(slSection, "DBPath", tmCsiFtpInfo.sLogPathName)
    tmCsiFtpInfo.sLogPathName = Trim$(tmCsiFtpInfo.sLogPathName) & "\" & "Messages\FTPLog.txt"
    ilRet = csiFTPInit(tmCsiFtpInfo)
   
        

    Exit Function
End Function

Private Function mFtpWebFileToServer(slFileName As String) As Boolean

    Dim ilRet As Integer
    Dim FTPIsOn As String
    Dim FTPInfo As CSIFTPINFO
    Dim FTPStatus As CSIFTPSTATUS
    Dim FTPErrorInfo As CSIFTPERRORINFO
    
    On Error GoTo ErrHand
    
    mFtpWebFileToServer = True
'    Exit Function
    
    mFtpWebFileToServer = False
    
    
    mFtpWebFileToServer = False
    ' First load all the information we need from the ini file.
    Call gLoadOption(sgWebServerSection, "FTPIsOn", FTPIsOn)
    If Val(FTPIsOn) < 1 Then
        ' FTP is turned off. Return success.
        ' Note: This will be the case when the affiliate system and IIS is running on the same machine.
        '       Usually only while testing.
        mFtpWebFileToServer = True
        Exit Function
    End If
    
    
    ' Receive the following file from the web server.
    ilRet = csiFTPFileFromServer(Trim$(slFileName))
    ilRet = csiFTPGetStatus(FTPStatus)
    While FTPStatus.iState = 1
        DoEvents
        Sleep (2000)
        ilRet = csiFTPGetStatus(FTPStatus)
    Wend
    If FTPStatus.iStatus <> 0 Then
        ' Errors occured.
        ilRet = csiFTPGetError(FTPErrorInfo)
        gMsgBox "FTP Failed. " & FTPErrorInfo.sInfo
        gMsgBox "The file name was " & FTPErrorInfo.sFileThatFailed
        Exit Function
    End If
    
    mFtpWebFileToServer = True
    
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    SetResults "Error: FTP", RGB(255, 0, 0)
    SetResults "See Error Log", RGB(255, 0, 0)
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebImportAiredSpot-mFtpWebFileToServer: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "WebImportLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    
    'Debug
    'Resume Next
End Function

Private Function mImportMG(llIdx As Long) As Boolean

    'D.S. 03/10/12
    
    Dim llAstCode As Long
    Dim ilIndex As Integer
    Dim llLstCode As Long
    Dim llTemp As Long
    Dim ilAdfCode As Integer
    Dim llDATCode As Long
    Dim llCpfCode As Long
    Dim llRsfCode As Long
    Dim llCntrNo As Long
    Dim ilLen As Integer
    Dim slStationCompliant As String
    Dim slAgencyCompliant As String
    Dim slAffidavitSource As String
    Dim llLkAstCode As Long
    Dim slSQLQuery As String
    Dim rst_Ast As ADODB.Recordset
    Dim blCpfAgree As Boolean

    On Error GoTo ErrHand
    
    mImportMG = False
    llLstCode = mAddLst(llIdx)
    If llLstCode = 0 Then
        gLogMsg "Warning: A MG Spot with an original spot code = " & tmAirSpotInfo(llIdx).lOrgAstCode & " was imported, but that spot code no longer exist", "WebImportLog.Txt", False
        mImportMG = True
        Exit Function
    End If
    
    llLkAstCode = tmAirSpotInfo(llIdx).lOrgAstCode
    slSQLQuery = "Select * FROM ast WHERE (AstCode = " & tmAirSpotInfo(llIdx).lOrgAstCode & ")"
    Set rst_Ast = gSQLSelectCall(slSQLQuery)
    
    If Not rst_Ast.EOF Then
        If rst_Ast!astLkAstCode > 0 Then
            gLogMsg "Warning: A Previously recived MG Spot AST code = " & rst_Ast!astLkAstCode & " was found linked to AST code = " & tmAirSpotInfo(llIdx).lOrgAstCode, "WebImportLog.Txt", False
            tmAirSpotInfo(llIdx).iFound = True
            tmAirSpotInfo(llIdx).iUpdateComplete = True
            mImportMG = True
            Exit Function
        End If
    End If
    
    If Not rst_Ast.EOF Then
        ilAdfCode = rst_Ast!astAdfCode
        llDATCode = rst_Ast!astDatCode
        llCpfCode = mGetCpfCode(ilAdfCode, Trim(tmAirSpotInfo(llIdx).sISCI))
        
        'D.S. 02/5/19 start new code
        blCpfAgree = False
        If llCpfCode = rst_Ast!astCpfCode Then
            blCpfAgree = True
        End If
        'D.S. 02/5/19 end new code
        
        llRsfCode = rst_Ast!astRsfCode
        llCntrNo = rst_Ast!astCntrNo
        ilLen = rst_Ast!astLen
    Else
        ilAdfCode = 0
        llDATCode = 0
        llCpfCode = 0
        llRsfCode = 0
        llCntrNo = 0
        ilLen = 0
    End If
    slStationCompliant = ""
    slAgencyCompliant = ""
    slAffidavitSource = tmAirSpotInfo(llIdx).sVendorSource
    slSQLQuery = "INSERT INTO ast"
    slSQLQuery = slSQLQuery & "(astCode, astAtfCode, astShfCode, astVefCode, "
    slSQLQuery = slSQLQuery & "astSdfCode, astLsfCode, astAirDate, astAirTime, "
    slSQLQuery = slSQLQuery & "astStatus, astCPStatus, astFeedDate, astFeedTime, "
    slSQLQuery = slSQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
    slSQLQuery = slSQLQuery & " VALUES "
    slSQLQuery = slSQLQuery & "(" & "Replace" & ", " & tmAirSpotInfo(llIdx).lAtfCode & ", " & tmAirSpotInfo(llIdx).iShfCode & ", "
    slSQLQuery = slSQLQuery & tmAirSpotInfo(llIdx).iVefCode & ", " & 0 & ", " & llLstCode & ", "
    slSQLQuery = slSQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', '" & Format$(tmAirSpotInfo(llIdx).sActualAirTime1, sgSQLTimeForm) & "', "
    'D.S. 01/31/19 start new code

'    If llCpfCode = 0 Then
'        slSQLQuery = slSQLQuery & ASTEXTENDED_MG + ASTEXTENDED_ISCICHGD & ", " & 1 & ", '" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', "
'    Else
'        slSQLQuery = slSQLQuery & ASTEXTENDED_MG & ", " & 1 & ", '" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', "
'    End If
    'D.S. 01/31/19 end new code
    
    'D.S. 02/5/19 start new code
    If blCpfAgree Then
        slSQLQuery = slSQLQuery & ASTEXTENDED_MG & ", " & 1 & ", '" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', "
    Else
        slSQLQuery = slSQLQuery & ASTEXTENDED_MG + ASTEXTENDED_ISCICHGD & ", " & 1 & ", '" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', "
    End If
    'D.S. 02/5/19 end new code
    
    slSQLQuery = slSQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sActualAirTime1, sgSQLTimeForm) & "', "
    slSQLQuery = slSQLQuery & ilAdfCode & ", " & llDATCode & ", " & llCpfCode & ", " & llRsfCode & ", "
    slSQLQuery = slSQLQuery & "'" & slStationCompliant & "', '" & slAgencyCompliant & "', '" & slAffidavitSource & "', " & llCntrNo & ", " & ilLen & ", " & llLkAstCode & ", " & 0 & ", " & igUstCode & ")"
    llAstCode = gInsertAndReturnCode(slSQLQuery, "ast", "astCode", "Replace")
    
    'Update the original AST record with the link to the new AST record that was created in the above statement
    slSQLQuery = "UPDATE ast SET "
    slSQLQuery = slSQLQuery + "astAffidavitSource = '" & slAffidavitSource & "', "
    slSQLQuery = slSQLQuery + "astLen = '" & tmAirSpotInfo(llIdx).iSpotLen & "', "
    slSQLQuery = slSQLQuery + "astLkAstCode = " & llAstCode
    slSQLQuery = slSQLQuery + " WHERE (astCode = " & tmAirSpotInfo(llIdx).lOrgAstCode & ")"
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportMG"
        mImportMG = False
        Exit Function
    End If
    'D.S. 01/31/19 start new code
    If llCpfCode = 0 Then
        llTemp = mAdjustISCIAsNeeded(llAstCode, ilAdfCode, rst_Ast!astCpfCode, tmAirSpotInfo(llIdx).sISCI, "M")
        'update the new MG ast rec with llTemp  astCpfCode
        slSQLQuery = "UPDATE ast SET "
        slSQLQuery = slSQLQuery + "astCpfCode = " & llTemp
        slSQLQuery = slSQLQuery + " WHERE (astCode = " & llAstCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            Screen.MousePointer = vbDefault
            gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportMG"
            mImportMG = False
            Exit Function
        End If
    End If
    'D.S. 01/31/19 end new code
    tmAirSpotInfo(llIdx).lAstCode = llAstCode
    tmAirSpotInfo(llIdx).iFound = True
    tmAirSpotInfo(llIdx).iUpdateComplete = True
    mImportMG = True
    rst_Ast.Close
    Exit Function
    
ErrHand:
    rst_Ast.Close
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mImportMG"
    mImportMG = False
    Exit Function
End Function

Private Function mImportBonus(llIdx As Long) As Boolean
    
    'D.S. 03/10/12
    
    'D.S. 03/10/12
    
    Dim llAstCode As Long
    Dim ilIndex As Integer
    Dim ilPdStatus As Integer
    Dim llLstCode As Long
    Dim ilAdfCode As Integer
    Dim llDATCode As Long
    Dim llCpfCode As Long
    Dim llRsfCode As Long
    Dim slStationCompliant As String
    Dim slAgencyCompliant As String
    Dim slAffidavitSource As String
    Dim llCntrNo As Long
    Dim ilLen As Integer

    On Error GoTo ErrHand
    
    mImportBonus = False
    
    
    llLstCode = mAddLst(llIdx)
    
    SQLQuery = "Select * FROM lst WHERE (lstCode = " & llLstCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilAdfCode = rst!lstAdfCode
        llDATCode = 0
        'llCpfCode = rst!lstCpfCode
        'llCpfCode = lmOrigCpfCode
        'D.S. 03/10/15
        llCpfCode = mGetCpfCode(ilAdfCode, Trim(tmAirSpotInfo(llIdx).sISCI))
        llRsfCode = 0
        llCntrNo = rst!lstCntrNo
        ilLen = rst!lstLen
    Else
        ilAdfCode = 0
        llDATCode = 0
        llCpfCode = 0
        llRsfCode = 0
        llCntrNo = 0
        ilLen = 0
    End If
    slStationCompliant = ""
    slAgencyCompliant = ""
    slAffidavitSource = tmAirSpotInfo(llIdx).sVendorSource
    SQLQuery = "INSERT INTO ast"
    SQLQuery = SQLQuery & "(astCode, astAtfCode, astShfCode, astVefCode, "
    SQLQuery = SQLQuery & "astSdfCode, astLsfCode, astAirDate, astAirTime, "
    '12/13/13: Support New AST layout
    'SQLQuery = SQLQuery & "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, "
    SQLQuery = SQLQuery & "astStatus, astCPStatus, astFeedDate, astFeedTime, "
    'SQLQuery = SQLQuery & "astPledgeStartTime, astPledgeEndTime, astPledgeStatus)"
    SQLQuery = SQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
    SQLQuery = SQLQuery & " VALUES "
    SQLQuery = SQLQuery & "(" & "Replace" & ", " & tmAirSpotInfo(llIdx).lAtfCode & ", " & tmAirSpotInfo(llIdx).iShfCode & ", "
    SQLQuery = SQLQuery & tmAirSpotInfo(llIdx).iVefCode & ", " & 0 & ", " & llLstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', '" & Format$(tmAirSpotInfo(llIdx).sActualAirTime1, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & ASTEXTENDED_BONUS & ", " & 1 & ", '" & Format$(tmAirSpotInfo(llIdx).sFeedDate, sgSQLDateForm) & "', "
    'SQLQuery = SQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sFeedTime, sgSQLTimeForm) & "', '" & Format$(tmAirSpotInfo(llIdx).sPledgeStartDate1, sgSQLDateForm) & "', "
    'SQLQuery = SQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sPledgeStartTime, sgSQLTimeForm) & "', '" & Format$(tmAirSpotInfo(llIdx).sPledgeEndTime, sgSQLTimeForm) & "', " & ASTEXTENDED_BONUS & ")"
    SQLQuery = SQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sFeedTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & ilAdfCode & ", " & llDATCode & ", " & llCpfCode & ", " & llRsfCode & ", "
    SQLQuery = SQLQuery & "'" & slStationCompliant & "', '" & slAgencyCompliant & "', '" & slAffidavitSource & "', " & llCntrNo & ", " & ilLen & ", " & 0 & ", " & 0 & ", " & igUstCode & ")"
    llAstCode = gInsertAndReturnCode(SQLQuery, "ast", "astCode", "Replace")
    
    tmAirSpotInfo(llIdx).lAstCode = llAstCode
    tmAirSpotInfo(llIdx).iFound = True
    tmAirSpotInfo(llIdx).iUpdateComplete = True

    mImportBonus = True
        
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mImportBonus"
    mImportBonus = False
    Exit Function
End Function

Private Function mImportMissedReason(llIdx As Long) As Boolean
    
    Dim llAstCode As Long
    Dim ilIndex As Integer
    Dim ilPdStatus As Integer
    
    On Error GoTo ErrHand
    
    mImportMissedReason = False
    
    If tmAirSpotInfo(llIdx).iMissedReason <= 0 Then
        mImportMissedReason = True
        Exit Function
    End If
    SQLQuery = "UPDATE ast SET "
    SQLQuery = SQLQuery + "astMissedMnfCode = " & tmAirSpotInfo(llIdx).iMissedReason
    SQLQuery = SQLQuery + " WHERE (astCode = " & tmAirSpotInfo(llIdx).lAstCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportMissedReason"
        mImportMissedReason = False
        Exit Function
    End If
        
    mImportMissedReason = True
        
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mImportMissedReason"
    mImportMissedReason = False
    Exit Function
End Function


Private Function mImportReplacement(llIdx As Long) As Boolean

    'D.S. 03/10/12
    
    Dim llAstCode As Long
    Dim ilIndex As Integer
    Dim ilPdStatus As Integer
    Dim llLstCode As Long
    Dim ilAdfCode As Integer
    Dim llDATCode As Long
    Dim llCpfCode As Long
    Dim llRsfCode As Long
    Dim llCntrNo As Long
    Dim ilLen As Integer
    Dim slStationCompliant As String
    Dim slAgencyCompliant As String
    Dim slAffidavitSource As String
    Dim llLkAstCode As Long

    On Error GoTo ErrHand
    
    mImportReplacement = False
    
    
    llLstCode = mAddLst(llIdx)
    
    SQLQuery = "Select * FROM lst WHERE (lstCode = " & llLstCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        ilAdfCode = rst!lstAdfCode
        llDATCode = 0
        llCpfCode = mGetCpfCode(ilAdfCode, Trim(tmAirSpotInfo(llIdx).sISCI))
        llRsfCode = 0
        llCntrNo = 0
        ilLen = rst!lstLen
    Else
        ilAdfCode = 0
        llDATCode = 0
        llCpfCode = 0
        llRsfCode = 0
        llCntrNo = 0
        ilLen = 0
    End If
    llLkAstCode = tmAirSpotInfo(llIdx).lOrgAstCode
    slStationCompliant = ""
    slAgencyCompliant = ""
    slAffidavitSource = tmAirSpotInfo(llIdx).sVendorSource
    'Inserting the replacement spot
    SQLQuery = "INSERT INTO ast"
    SQLQuery = SQLQuery & "(astCode, astAtfCode, astShfCode, astVefCode, "
    SQLQuery = SQLQuery & "astSdfCode, astLsfCode, astAirDate, astAirTime, "
    SQLQuery = SQLQuery & "astStatus, astCPStatus, astFeedDate, astFeedTime, "
    SQLQuery = SQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
    SQLQuery = SQLQuery & " VALUES "
    SQLQuery = SQLQuery & "(" & "Replace" & ", " & tmAirSpotInfo(llIdx).lAtfCode & ", " & tmAirSpotInfo(llIdx).iShfCode & ", "
    SQLQuery = SQLQuery & tmAirSpotInfo(llIdx).iVefCode & ", " & 0 & ", " & llLstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', '" & Format$(tmAirSpotInfo(llIdx).sActualAirTime1, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & ASTEXTENDED_REPLACEMENT & ", " & 1 & ", '" & Format$(tmAirSpotInfo(llIdx).sFeedDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sFeedTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & ilAdfCode & ", " & llDATCode & ", " & llCpfCode & ", " & llRsfCode & ", "
    SQLQuery = SQLQuery & "'" & slStationCompliant & "', '" & slAgencyCompliant & "', '" & slAffidavitSource & "', " & llCntrNo & ", " & ilLen & ", " & llLkAstCode & ", " & 0 & ", " & igUstCode & ")"
    llAstCode = gInsertAndReturnCode(SQLQuery, "ast", "astCode", "Replace")
  
    'Doug: why are you updating astStatus?
    'Are you assuming that it was not marked as missed yet?
    'You received the replacement prior to Missed
    'Why are you not updating the same status in ImportMG?
    SQLQuery = "UPDATE ast SET "
    SQLQuery = SQLQuery + "astAffidavitSource = '" & Trim(slAffidavitSource) & "', "
    SQLQuery = SQLQuery + "astLkAstCode = " & llAstCode & ","
    SQLQuery = SQLQuery + "astStatus = " & 4 & ","
    SQLQuery = SQLQuery + "astCPStatus = " & 1
    SQLQuery = SQLQuery + " WHERE (astCode = " & tmAirSpotInfo(llIdx).lOrgAstCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "WebImportLog.txt", "WebImportSchdSpot-mImportReplacement"
        mImportReplacement = False
        Exit Function
    End If
    tmAirSpotInfo(llIdx).lAstCode = llAstCode
    tmAirSpotInfo(llIdx).iFound = True
    tmAirSpotInfo(llIdx).iUpdateComplete = True
        
    mImportReplacement = True
        
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mImportReplacement"
    mImportReplacement = False
    Exit Function
End Function

Private Function mAddLst(llIdx As Long) As Long
    
    Dim slProd As String
    Dim slCart As String
    Dim slISCI As String
    Dim ilAdf As Integer
    Dim ilAdfCode As Integer
    Dim llCntrNo As Long
    Dim llLineVef As Long
    Dim ilLineNo As Integer
    Dim ilAgfCode As Integer
    Dim ilPriceType As Integer
    Dim ilStatus As Integer
    Dim ilLoop As Integer
    Dim llLst As Long
    Dim llSdfCode As Long
    Dim llTemp As Long
    Dim rst_Temp As ADODB.Recordset
    ReDim ilDay(0 To 6) As Integer
    Dim llGsfCode As Long
    
    mAddLst = 0
    On Error GoTo ErrHand
    
    SQLQuery = "SELECT astLsfCode, astCpfCode FROM ast where astCode = " & tmAirSpotInfo(llIdx).lOrgAstCode
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    If Not rst_Temp.EOF Then
        llTemp = rst_Temp!astLsfCode
        lmOrigCpfCode = rst_Temp!astCpfCode
    Else
        'No matching original ast code found.  Must have been deleted
        tmAirSpotInfo(llIdx).iFound = 0
        mAddLst = 0
        Exit Function
    End If
    
    SQLQuery = "SELECT * FROM Lst where lstCode = " & llTemp
    Set rst_Temp = gSQLSelectCall(SQLQuery)
    If rst_Temp.EOF Then
        'Error condition
    End If
    
    slProd = tmAirSpotInfo(llIdx).sProd
    slCart = tmAirSpotInfo(llIdx).sCart
    slISCI = gFixQuote(tmAirSpotInfo(llIdx).sISCI)
    ilAdfCode = -1
    For ilAdf = 0 To UBound(tgAdvtInfo) - 1 Step 1
        If StrComp(UCase(Trim$(tgAdvtInfo(ilAdf).sAdvtName)), UCase(Trim$(tmAirSpotInfo(llIdx).sAdvt)), vbTextCompare) = 0 Then
            ilAdfCode = tgAdvtInfo(ilAdf).iCode
            Exit For
        End If
    Next ilAdf
    '7/7/20: The above code is failing if the name had an apostrophe (it is returned as two single quotes)
    If ilAdfCode = -1 Then
        ilAdfCode = rst_Temp!lstAdfCode
    End If
    llCntrNo = rst_Temp!lstCntrNo
    ilLineNo = rst_Temp!lstLineNo
    ilAgfCode = rst_Temp!lstAgfCode
    ilPriceType = rst_Temp!lstPriceType
    llLineVef = rst_Temp!lstLnVefCode
    llSdfCode = rst_Temp!lstSdfCode
    If tmAirSpotInfo(llIdx).lgsfCode <> -1 Then
        llGsfCode = tmAirSpotInfo(llIdx).lgsfCode
    Else
        llGsfCode = rst_Temp!lstGsfCode
    End If
    
    For ilLoop = 0 To 6 Step 1
        ilDay(ilLoop) = 0
    Next ilLoop
    If Trim(tmAirSpotInfo(llIdx).sRecType) = "M" Then
        ilStatus = ASTEXTENDED_MG
    ElseIf Trim(tmAirSpotInfo(llIdx).sRecType) = "B" Then
        ilStatus = ASTEXTENDED_BONUS
    ElseIf Trim(tmAirSpotInfo(llIdx).sRecType) = "R" Then
        ilStatus = ASTEXTENDED_REPLACEMENT
    End If
    
    ilDay(gWeekDayLong(gDateValue(tmAirSpotInfo(llIdx).sActualAirDate1))) = 1
    SQLQuery = "INSERT INTO lst (lstCode, lstType, lstSdfCode, lstCntrNo, "
    SQLQuery = SQLQuery & "lstAdfCode, lstAgfCode, lstProd, "
    SQLQuery = SQLQuery & "lstLineNo, lstLnVefCode, lstStartDate, "
    SQLQuery = SQLQuery & "lstEndDate, lstMon, lstTue, "
    SQLQuery = SQLQuery & "lstWed, lstThu, lstFri, "
    SQLQuery = SQLQuery & "lstSat, lstSun, lstSpotsWk, "
    SQLQuery = SQLQuery & "lstPriceType, lstPrice, lstSpotType, "
    SQLQuery = SQLQuery & "lstLogVefCode, lstLogDate, lstLogTime, "
    SQLQuery = SQLQuery & "lstDemo, lstAud, lstISCI, "
    SQLQuery = SQLQuery & "lstWkNo, lstBreakNo, lstPositionNo, "
    SQLQuery = SQLQuery & "lstSeqNo, lstZone, lstCart, "
    SQLQuery = SQLQuery & "lstCpfCode, lstCrfCsfCode, lstStatus, "
    SQLQuery = SQLQuery & "lstLen, lstUnits, lstCifCode, "
    SQLQuery = SQLQuery & "lstAnfCode, lstEvtIDCefCode, lstSplitNetwork, "
    'SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, lstUnused)"
    SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, "
    SQLQuery = SQLQuery & "lstLnStartTime, lstLnEndTime, lstUnused)"
    SQLQuery = SQLQuery & " VALUES (" & "Replace" & ", " & 2 & ", " & llSdfCode & ", " & llCntrNo & ", "
    SQLQuery = SQLQuery & ilAdfCode & ", " & ilAgfCode & ", '" & slProd & "', "
    
    'SQLQuery = SQLQuery & ilLineNo & ", " & tmAirSpotInfo(llIdx).iVefCode & ", '" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & ilLineNo & ", " & llLineVef & ", '" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', " & ilDay(0) & ", " & ilDay(1) & ", "
    SQLQuery = SQLQuery & ilDay(2) & ", " & ilDay(3) & ", " & ilDay(4) & ", "
    SQLQuery = SQLQuery & ilDay(5) & ", " & ilDay(6) & ", " & 0 & ", "
    SQLQuery = SQLQuery & ilPriceType & ", " & 0 & ", " & 5 & ", "
    SQLQuery = SQLQuery & tmAirSpotInfo(llIdx).iVefCode & ", '" & Format$(tmAirSpotInfo(llIdx).sActualAirDate1, sgSQLDateForm) & "', '" & Format$(tmAirSpotInfo(llIdx).sActualAirTime1, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & 0 & "', " & 0 & ", '" & slISCI & "', "
    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", '" & "" & "', '" & slCart & "', "
    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & ilStatus & ", "
    SQLQuery = SQLQuery & tmAirSpotInfo(llIdx).iSpotLen & ", " & 0 & ", " & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", " & 0 & ", '" & "N" & "', "
    'SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", "
    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & llGsfCode & ", '" & "N" & "', " & 0 & ", "
    SQLQuery = SQLQuery & "'" & Format("12am", sgSQLTimeForm) & "', '" & Format("12am", sgSQLTimeForm) & "', '" & "" & "'" & ")"
    llLst = gInsertAndReturnCode(SQLQuery, "lst", "lstCode", "Replace")
    rst_Temp.Close
    
    mAddLst = llLst
    
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebImportAiredSpot-mAddList"
    On Error GoTo 0
    mAddLst = False
    Exit Function
End Function

Private Sub mLogTimingResults()

    'D.S 02/26/13
    Exit Sub

    Dim llTime As Long
    
    lgETime25 = timeGetTime
    lgTtlTime25 = lgTtlTime25 + lgETime25 - lgSTime25
    
    
    gLogMsg "*** mInitiateExport - mExport time  = " & gTimeString(lgTtlTime24 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "*** cmdExport click = " & gTimeString(lgTtlTime24 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "*** mExportSpots = " & gTimeString(lgTtlTime23 / 1000, True), "WebExpSummary.Txt", False
    llTime = lgTtlTime4 + lgTtlTime2 + lgTtlTime5
    llTime = lgTtlTime23 - llTime
'    gLogMsg "     mExportSpots without ggAstInfo, BuildDetail Recs or buildHeaders  = " & gTimeString(llTime / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "    Total Import Time = " & gTimeString(lgTtlTime5 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "    BuildDetail Recs = " & gTimeString(lgTtlTime4 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       mEstimateDateAndTime " & lgCount11, "WebExpSummary.Txt", False
    gLogMsg "       ObtainRotEndDete = " & gTimeString(lgTtlTime15 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       Number of Calls to ObtainRotEndDete = " & CStr(lgCount8), "WebExpSummary.Txt", False
    gLogMsg "       Format slRotCpyEndDate = " & gTimeString(lgTtlTime19 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       GetCFS = " & gTimeString(lgTtlTime16 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       Build Game Inf = " & gTimeString(lgTtlTime17 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       Print string to File = " & gTimeString(lgTtlTime18 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "    GGetAstInfo Time = " & gTimeString(lgTtlTime2 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "       Regional Copy Call = " & gTimeString(lgTtlTime8 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "           RC was called " & lgCount1 & " times out of a possible " & lgCount2 & " times.", "WebExpSummary.Txt", False
    'gLogMsg "            # Blackouts Unchanged = " & lgCount6 & "  # Blackout Added = " & lgCount5 & "  # Blackout Updated = " & lgCount4, "WebExpSummary.Txt", False
    'gLogMsg "            # Blackouts Unchanged = " & lgCount6 & "  # Blackout Added = " & lgCount5 & "  # Blackout Updated = " & lgCount4 & "  # Blackout Removed = " & lgCount7, "WebExpSummary.Txt", False
    'gLogMsg "            Time to validate Blackout Ok = " & gTimeString(lgTtlTime9 / 1000, True), "WebExpSummary.Txt", False
    'gLogMsg "            Time to add Blackout = " & gTimeString(lgTtlTime10 / 1000, True), "WebExpSummary.Txt", False
    'gLogMsg "            Time to update Blackout = " & gTimeString(lgTtlTime11 / 1000, True), "WebExpSummary.Txt", False
    'gLogMsg "            Time to remove Blackout = " & gTimeString(lgTtlTime12 / 1000, True), "WebExpSummary.Txt", False
    'gLogMsg "            Time checking if Blackout exist and should be removed but does not exist = " & lgTtlTime13 / 1000, "WebExpSummary.Txt", False
    gLogMsg "           gSeparateRegions+gRegionTestDefinition = " & gTimeString(lgTtlTime14 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "           Binary Search " & gTimeString(lgTtlTime20 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "           Get Copy " & gTimeString(lgTtlTime21 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "               Get Copy Found in Array " & lgCount9, "WebExpSummary.Txt", False
    gLogMsg "               Get Copy NOT Found in Array " & lgCount10, "WebExpSummary.Txt", False
    gLogMsg "               SQL in Get Copy " & gTimeString(lgTtlTime22 / 1000, True), "WebExpSummary.Txt", False
    
    
    gLogMsg "*** FTP Time = " & gTimeString(lgTtlTime3 / 1000, True), "WebExpSummary.Txt", False
    gLogMsg "*** Web Import Processing Time = " & gTimeString(lgTtlTime6 / 1000, True), "WebExpSummary.Txt", False
    
    lgETime1 = timeGetTime
    lgTtlTime1 = lgETime1 - lgSTime1
    gLogMsg "Total Export Time = " & gTimeString(lgTtlTime1 / 1000, True), "WebImpSummary.Txt", False
    'gLogMsg " ", "WebExpSummary.Txt", False

End Sub


Private Function mGetPledgeByEvent() As String
    
    Dim ilVff As Integer
    
    mGetPledgeByEvent = "N"
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        Exit Function
    End If
    If imVefCode <= 0 Then
        Exit Function
    End If
    ilVff = gBinarySearchVff(imVefCode)
    If ilVff <> -1 Then
        If Trim$(tgVffInfo(ilVff).sPledgeByEvent) = "" Then
            mGetPledgeByEvent = "N"
        Else
            mGetPledgeByEvent = Trim$(tgVffInfo(ilVff).sPledgeByEvent)
        End If
    End If
End Function

Private Function mGetSports()

    'D.S. 02/26/13 Needed this information to support auto-import

    Dim rst As ADODB.Recordset

    SQLQuery = "SELECT spfSportInfo"
    SQLQuery = SQLQuery + " FROM SPF_Site_Options"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If IsNull(rst!spfSportInfo) Or (Len(rst!spfSportInfo) = 0) Then
            sgSpfSportInfo = Chr$(0)
        Else
            sgSpfSportInfo = rst!spfSportInfo
        End If
    Else
        sgSpfSportInfo = Chr$(0)
    End If
    
    rst.Close


End Function

Private Function mInsertCpfIntoAst(lAstCode As Long) As Long



End Function



Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "ASI"
End Sub

Private Function mGetCpfCode(iAdvCode As Integer, sISCI As String) As Long

    'D.S. Created 03/10/15
    'Determine the CPFCode based off of the advertiser code and the ISCI code
    Dim rst As ADODB.Recordset
    Dim ilAdvCode As Integer
    Dim ilIdx As Integer
    Dim slTempStr As String
     
    mGetCpfCode = 0
    'D.S. 11/14/19 TTP 9617 Added two lines below to handle embedded single quotes in the ISCI string.
    slTempStr = sISCI
    sISCI = gFixQuote(slTempStr)
    SQLQuery = "select distinct cpfcode"
    SQLQuery = SQLQuery + " from CIF_Copy_Inventory, CPF_Copy_Prodct_ISCI"
    SQLQuery = SQLQuery + " where cifadfcode = " & iAdvCode & " and cpfisci = " & "'" & sISCI & "'"
    SQLQuery = SQLQuery + " And cifCpfCode = cpfCode"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        mGetCpfCode = rst!cpfCode
    End If
    rst.Close
End Function

Private Sub mReImport()
    Dim ilRet As Integer
    
    'ilRet = gPopAttInfo()
    'ilRet = gPopAll
    Set myEnt = New CENThelper
    With myEnt
        .TypeEnt = Receivedpostedfromweb
        .User = igUstCode
        .ThirdParty = Web
        .ErrorLog = cmPathForgLogMsg
        .ProcessStart
    End With
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    Screen.MousePointer = vbHourglass
    lmTotalHeaders = 0
    lmTotalSpots = 0
    lmTotalActivityLogs = 0
    imImporting = True
    smWebSpots = "WebSpots_" & "ReImport.txt"
    ilRet = mProcessSpotFile()
    If ilRet Then
        sgReImportStatus = "Counterpoint Affidavit: Import Successful"
    End If
    Set myEnt = Nothing
    imImporting = False
    tmcTerminate.Enabled = True
End Sub

Private Sub tmcTerminate_Timer()
    tmcTerminate.Enabled = False
    Screen.MousePointer = vbDefault
    Unload frmWebImportAiredSpot
End Sub

Private Sub mCheckForMG(tlAstInfo As ASTINFO)
    Dim llMGLstCode As Long
    Dim llMGAstCode As Long
    Dim slPdDate As String
    Dim slPdTime As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim llLstCode As Long
    Dim llAstCode As Long
    Dim slAffidavitSource As String
    Dim ilIdx As Integer
    Dim slSource As String
    
    If (tlAstInfo.iStatus Mod 100 <> ASTAIR_NA_OTHER) And (tlAstInfo.iStatus Mod 100 <> ASTAIR_NOTCARRIED) And (tlAstInfo.iStatus Mod 100 <> ASTAIR_MISSED_MG_BYPASS) And (tlAstInfo.iStatus Mod 100 <= ASTAIR_CMML) Then
        slPdDate = tlAstInfo.sPledgeDate
        slPdTime = tlAstInfo.sPledgeStartTime
        slAirDate = tlAstInfo.sAirDate
        slAirTime = tlAstInfo.sAirTime
        llLstCode = tlAstInfo.lLstCode
        llAstCode = tlAstInfo.lCode
        If gObtainPrevMonday(slPdDate) <> gObtainPrevMonday(slAirDate) Then
            'Create the MG spots
            llMGLstCode = mAddMGLst(llLstCode, slAirDate, slAirTime)
            If llMGLstCode > 0 Then
                slSource = "UD"
                For ilIdx = 0 To UBound(tmAirSpotInfo)
                    If tmAirSpotInfo(ilIdx).lAstCode = llAstCode Then
                        slSource = tmAirSpotInfo(ilIdx).sVendorSource
                        Exit For
                    End If
                Next ilIdx
                slAffidavitSource = slSource
                llMGAstCode = mAddAstMG(llMGLstCode, tlAstInfo, slAffidavitSource)
                If llMGAstCode > 0 Then
                    tlAstInfo.iStatus = 4
                    tlAstInfo.sAirDate = slPdDate
                    tlAstInfo.sAirTime = slPdTime
                    mChgAstToMissed llMGAstCode, tlAstInfo
                Else
                End If
            Else
            End If
        End If
    End If
End Sub

Private Function mAddMGLst(llLsfCode As Long, slAirDate As String, slAirTime As String) As Long
    Dim llLst As Long
    
    On Error GoTo ErrHand
    SQLQuery = "SELECT * From lst Where lstCode = " & llLsfCode
    Set lst_rst = gSQLSelectCall(SQLQuery)
    If lst_rst.EOF Then
        mAddMGLst = 0
        Exit Function
    End If
    
    SQLQuery = "Insert Into lst ( "
    SQLQuery = SQLQuery & "lstCode, "
    SQLQuery = SQLQuery & "lstType, "
    SQLQuery = SQLQuery & "lstSdfCode, "
    SQLQuery = SQLQuery & "lstCntrNo, "
    SQLQuery = SQLQuery & "lstAdfCode, "
    SQLQuery = SQLQuery & "lstAgfCode, "
    SQLQuery = SQLQuery & "lstProd, "
    SQLQuery = SQLQuery & "lstLineNo, "
    SQLQuery = SQLQuery & "lstLnVefCode, "
    SQLQuery = SQLQuery & "lstStartDate, "
    SQLQuery = SQLQuery & "lstEndDate, "
    SQLQuery = SQLQuery & "lstMon, "
    SQLQuery = SQLQuery & "lstTue, "
    SQLQuery = SQLQuery & "lstWed, "
    SQLQuery = SQLQuery & "lstThu, "
    SQLQuery = SQLQuery & "lstFri, "
    SQLQuery = SQLQuery & "lstSat, "
    SQLQuery = SQLQuery & "lstSun, "
    SQLQuery = SQLQuery & "lstSpotsWk, "
    SQLQuery = SQLQuery & "lstPriceType, "
    SQLQuery = SQLQuery & "lstPrice, "
    SQLQuery = SQLQuery & "lstSpotType, "
    SQLQuery = SQLQuery & "lstLogVefCode, "
    SQLQuery = SQLQuery & "lstLogDate, "
    SQLQuery = SQLQuery & "lstLogTime, "
    SQLQuery = SQLQuery & "lstDemo, "
    SQLQuery = SQLQuery & "lstAud, "
    SQLQuery = SQLQuery & "lstISCI, "
    SQLQuery = SQLQuery & "lstWkNo, "
    SQLQuery = SQLQuery & "lstBreakNo, "
    SQLQuery = SQLQuery & "lstPositionNo, "
    SQLQuery = SQLQuery & "lstSeqNo, "
    SQLQuery = SQLQuery & "lstZone, "
    SQLQuery = SQLQuery & "lstCart, "
    SQLQuery = SQLQuery & "lstCpfCode, "
    SQLQuery = SQLQuery & "lstCrfCsfCode, "
    SQLQuery = SQLQuery & "lstStatus, "
    SQLQuery = SQLQuery & "lstLen, "
    SQLQuery = SQLQuery & "lstUnits, "
    SQLQuery = SQLQuery & "lstCifCode, "
    SQLQuery = SQLQuery & "lstAnfCode, "
    SQLQuery = SQLQuery & "lstEvtIDCefCode, "
    SQLQuery = SQLQuery & "lstSplitNetwork, "
    SQLQuery = SQLQuery & "lstRafCode, "
    SQLQuery = SQLQuery & "lstFsfCode, "
    SQLQuery = SQLQuery & "lstGsfCode, "
    SQLQuery = SQLQuery & "lstImportedSpot, "
    SQLQuery = SQLQuery & "lstBkoutLstCode, "
    SQLQuery = SQLQuery & "lstLnStartTime, "
    SQLQuery = SQLQuery & "lstLnEndTime, "
    SQLQuery = SQLQuery & "lstUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & 2 & ", "    'lstType
    SQLQuery = SQLQuery & lst_rst!lstSdfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstCntrNo & ", "
    SQLQuery = SQLQuery & lst_rst!lstAdfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstAgfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstProd) & "', "
    SQLQuery = SQLQuery & lst_rst!lstLineNo & ", "
    SQLQuery = SQLQuery & lst_rst!lstLnVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(slAirDate, sgSQLDateForm) & "', "      'lstStartDate
    SQLQuery = SQLQuery & "'" & Format$(slAirDate, sgSQLDateForm) & "', "      'lstEndDate
    SQLQuery = SQLQuery & 0 & ", " 'lstMon
    SQLQuery = SQLQuery & 0 & ", " 'lstTue
    SQLQuery = SQLQuery & 0 & ", " 'lstWed
    SQLQuery = SQLQuery & 0 & ", " 'lstThu
    SQLQuery = SQLQuery & 0 & ", " 'lstFri
    SQLQuery = SQLQuery & 0 & ", " 'lstSat
    SQLQuery = SQLQuery & 0 & ", " 'lstSun
    SQLQuery = SQLQuery & 0 & ", " 'lstSpotsWk
    SQLQuery = SQLQuery & lst_rst!lstPriceType & ", "   'lstPriceType
    SQLQuery = SQLQuery & 0 & ", "   'lstPrice
    SQLQuery = SQLQuery & 5 & ", "    'lstSpotType
    SQLQuery = SQLQuery & lst_rst!lstLogVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(slAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote("0") & "', "  'lstDemo
    SQLQuery = SQLQuery & 0 & ", " 'lstAud
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstISCI) & "', "
    SQLQuery = SQLQuery & 0 & ", "    'lstWkNo
    SQLQuery = SQLQuery & 0 & ", " 'lstBreakNo
    SQLQuery = SQLQuery & 0 & ", "  'lstPositionNo
    SQLQuery = SQLQuery & 0 & ", "   'lstSeqNo
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstZone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(lst_rst!lstCart) & "', "
    SQLQuery = SQLQuery & lst_rst!lstCpfCode & ", "
    SQLQuery = SQLQuery & lst_rst!lstCrfCsfCode & ", "
    SQLQuery = SQLQuery & ASTEXTENDED_MG & ", "  'lstStatus
    SQLQuery = SQLQuery & lst_rst!lstLen & ", "
    SQLQuery = SQLQuery & 0 & ", "   'lstUnit
    SQLQuery = SQLQuery & lst_rst!lstCifCode & ", "
    SQLQuery = SQLQuery & 0 & ", " 'lstAnfCode
    SQLQuery = SQLQuery & 0 & ", "    'lstEvtIDCefCode
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'lstSplitNetwork
    SQLQuery = SQLQuery & 0 & ", "     'lstRafCode
    SQLQuery = SQLQuery & 0 & ", " 'lstFsfCode
    'Somewhere along the line we need to use the air date to determine the game number
    SQLQuery = SQLQuery & 0 & ", " 'lstGsfCode
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'lstImportedSpot
    SQLQuery = SQLQuery & 0 & ", "    'lstBkoutLstCode
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "  'lstLnStartTime
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "    'lstLnEndTime
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    llLst = gInsertAndReturnCode(SQLQuery, "lst", "lstCode", "Replace")
    If llLst > 0 Then
        mAddMGLst = llLst
    Else
        mAddMGLst = 0
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mChgAstToMissed"
End Function

Private Function mAddAstMG(llLstCode As Long, tlAstInfo As ASTINFO, slAffidavitSource As String) As Long
    Dim llAst As Long

    On Error GoTo ErrHand
    SQLQuery = "Insert Into ast ( "
    SQLQuery = SQLQuery & "astCode, "
    SQLQuery = SQLQuery & "astAtfCode, "
    SQLQuery = SQLQuery & "astShfCode, "
    SQLQuery = SQLQuery & "astVefCode, "
    SQLQuery = SQLQuery & "astSdfCode, "
    SQLQuery = SQLQuery & "astLsfCode, "
    SQLQuery = SQLQuery & "astAirDate, "
    SQLQuery = SQLQuery & "astAirTime, "
    SQLQuery = SQLQuery & "astStatus, "
    SQLQuery = SQLQuery & "astCPStatus, "
    SQLQuery = SQLQuery & "astFeedDate, "
    SQLQuery = SQLQuery & "astFeedTime, "
    SQLQuery = SQLQuery & "astAdfCode, "
    SQLQuery = SQLQuery & "astDatCode, "
    SQLQuery = SQLQuery & "astCpfCode, "
    SQLQuery = SQLQuery & "astRsfCode, "
    SQLQuery = SQLQuery & "astStationCompliant, "
    SQLQuery = SQLQuery & "astAgencyCompliant, "
    SQLQuery = SQLQuery & "astAffidavitSource, "
    SQLQuery = SQLQuery & "astCntrNo, "
    SQLQuery = SQLQuery & "astLen, "
    SQLQuery = SQLQuery & "astLkAstCode, "
    SQLQuery = SQLQuery & "astMissedMnfCode, "
    SQLQuery = SQLQuery & "astUstCode "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & tlAstInfo.lAttCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.iShttCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.iVefCode & ", "
    SQLQuery = SQLQuery & 0 & ", "         'astsdfCode
    SQLQuery = SQLQuery & llLstCode & ", "  'astlsfCode
    SQLQuery = SQLQuery & "'" & Format$(tlAstInfo.sAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAstInfo.sAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & ASTEXTENDED_MG & ", " 'astStatus
    SQLQuery = SQLQuery & tlAstInfo.iCPStatus & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlAstInfo.sAirDate, sgSQLDateForm) & "', "  'astFeedDate
    SQLQuery = SQLQuery & "'" & Format$(tlAstInfo.sAirTime, sgSQLTimeForm) & "', "  'astFeedTime
    SQLQuery = SQLQuery & tlAstInfo.iAdfCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.lDatCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.lCpfCode & ", "
    SQLQuery = SQLQuery & tlAstInfo.lRRsfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "  'astStationCompliant
    SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "   'astAgencyCompliant
    SQLQuery = SQLQuery & "'" & gFixQuote(slAffidavitSource) & "', "
    SQLQuery = SQLQuery & tlAstInfo.lCntrNo & ", "
    SQLQuery = SQLQuery & tlAstInfo.iLen & ", "
    SQLQuery = SQLQuery & tlAstInfo.lCode & ", "    'astLkAstCode
    SQLQuery = SQLQuery & 0 & ", "   'astMissedMnfCode
    SQLQuery = SQLQuery & igUstCode 'astUstCode
    SQLQuery = SQLQuery & ") "
    llAst = gInsertAndReturnCode(SQLQuery, "ast", "astCode", "Replace")
    If llAst > 0 Then
        mAddAstMG = llAst
    Else
        mAddAstMG = 0
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mChgAstToMissed"
End Function


Private Sub mChgAstToMissed(llMGAstCode As Long, tlAstInfo As ASTINFO)
    On Error GoTo ErrHand
    SQLQuery = "UPDATE ast SET "
    SQLQuery = SQLQuery & "astAirDate = '" & Format$(tlAstInfo.sAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "astAirTime = '" & Format$(tlAstInfo.sAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery + "astLkAstCode = " & llMGAstCode & ", "
    SQLQuery = SQLQuery + "astAgencyCompliant = '" & "Y" & "',"
    SQLQuery = SQLQuery + "astStationCompliant = '" & "Y" & "',"
    SQLQuery = SQLQuery + "astStatus = " & tlAstInfo.iStatus
    SQLQuery = SQLQuery + " WHERE (astCode = " & tlAstInfo.lCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand1:
        Screen.MousePointer = vbDefault
        gHandleError "WebImportLog.txt", "WebImportSchdSpot-mChgAstToMissed"
        Exit Sub
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Set MG-mChgAstToMissed"
'ErrHand1:
'    gHandleError "AffErrorLog.txt", "Set MG-mChgAstToMissed"
'    Return
End Sub

Private Function mAdjustISCIAsNeeded(lAstCode As Long, iAdfCode As Integer, lCpfCode As Long, sImportISCI As String, sRecType As String) As Long
    '06/16/16 This function has been modified modified by Dick and Doug
    'return CpfCode if import isci is different; otherwise 0
    'add or update cpf with new isci; then create alt for old isci
    Dim slAstISCI As String
    Dim llRet As Long
    Dim rstISCI As ADODB.Recordset
    
    llRet = lCpfCode
    'Dan M 7/29/15 don't do 7639 until later: reports need to be fixed. To restore: lose goto line below
    '8018 comment out
   ' GoTo Cleanup
On Error GoTo ERRORBOX
    mAdjustISCIAsNeeded = llRet
    SQLQuery = "Select cpfISCI from CPF_Copy_Prodct_ISCI where cpfCode = " & lCpfCode
    Set rstISCI = gSQLSelectCall(SQLQuery)
    If rstISCI.EOF Then
        SetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
        'myErrors.WriteWarning "sql call invalid:" & SQLQuery, False
        gLogMsg "sql call invalid: " & SQLQuery, cmPathForgLogMsg, False
        Exit Function
    End If
    slAstISCI = Trim(rstISCI!cpfISCI)
    If slAstISCI <> sImportISCI Then
        SQLQuery = "select cpfCode from cpf_Copy_Prodct_ISCI where cpfisci = '" & sImportISCI & "'"
        Set rstISCI = gSQLSelectCall(SQLQuery)
        If Not rstISCI.EOF Then
            llRet = rstISCI!cpfCode
        Else
            SQLQuery = "INSERT into cpf_Copy_Prodct_ISCI (cpfCode,cpfName,cpfIsci,cpfCreative,cpfRotEndDate,cpfsifCode) VALUES (Replace,'','" & sImportISCI & "','','" & NODATE & "',0)"
                llRet = gInsertAndReturnCode(SQLQuery, "cpf_Copy_Prodct_ISCI", "cpfCode", "Replace")
            End If
            'now create alt
            If llRet <> 0 Then
                If sRecType <> "M" Then
                    mAddAltForIsci lAstCode, iAdfCode, lCpfCode
                End If
            Else
                SetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
               ' myErrors.WriteWarning "sql call invalid:" & SQLQuery, False
                gLogMsg "sql call invalid: " & SQLQuery, cmPathForgLogMsg, False
            End If
    rstISCI.Close
    End If
Cleanup:
    mAdjustISCIAsNeeded = llRet
    Exit Function
ERRORBOX:
    SetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
    'gHandleError smPathForgLogMsg, FORMNAME & "-mAdjustISCIAsNeeded"
    llRet = 0
    GoTo Cleanup
End Function
Private Sub mAddAltForIsci(llAstCode As Long, ilAdfCode As Integer, llCpfCode As Long)
    'note that the tmastInfo has never been updated with any new isci data
    SQLQuery = "insert into alt (altAstCode,altMissedDate,altAdfCode,altMgDate,altCpfCode) values (" & llAstCode & ",'" & NODATE & "'," & ilAdfCode & ",'" & NODATE & "'," & llCpfCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub errbox:
        SetResults "warning!  Issue in mAddAltForIsci. Please see log", MESSAGERED
        Exit Sub
    End If
    Exit Sub
errbox:
    SetResults "warning!  Issue in mAddAltForIsci. Please see log", MESSAGERED
    'gHandleError smPathForgLogMsg, FORMNAME & "-mAddAltForIsci"
End Sub

Public Function mCheckAnyAired(llAttCode As Long, slMoDate As String, slSuDate As String) As Boolean

    Dim ilStatus As Integer
    Dim ilLoop As Integer
    Dim rstAst As ADODB.Recordset
    Dim SQLQuery1 As String
    
    mCheckAnyAired = True
    SQLQuery1 = "Select count(*) as totalCount from ast "
    SQLQuery1 = SQLQuery1 + " WHERE (astAtfCode = " & llAttCode
    SQLQuery1 = SQLQuery1 + " AND Mod( astStatus, " & 100 & ") IN(0,1,6,7,9,10) "
    SQLQuery1 = SQLQuery1 + " AND (astFeedDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
    Set rstAst = gSQLSelectCall(SQLQuery1)
    If Not rstAst.EOF Then
       If rstAst!totalCount <= 0 Then
            mCheckAnyAired = False
        End If
    End If
    rstAst.Close
Exit Function


End Function

Sub mUpdateShttTables(ilShttCode As Integer, slWebPW As String, slStationEmail As String)
    '11/26/17
    Dim ilIndex As Integer
    Dim blRepopRequired As Boolean
    Dim slCallLetters As String
    
    blRepopRequired = False
    ilIndex = gBinarySearchStationInfoByCode(ilShttCode)
    If ilIndex <> -1 Then
        tgStationInfoByCode(ilIndex).sWebPW = Trim(slWebPW)
        tgStationInfoByCode(ilIndex).sWebEMail = Trim(slStationEmail)
        slCallLetters = Trim$(tgStationInfoByCode(ilIndex).sCallLetters)
        ilIndex = gBinarySearchStation(slCallLetters)
        If ilIndex <> -1 Then
            tgStationInfo(ilIndex).sWebPW = Trim(slWebPW)
            tgStationInfo(ilIndex).sWebEMail = Trim(slStationEmail)
        Else
            blRepopRequired = True
        End If
    Else
        blRepopRequired = True
    End If
    gFileChgdUpdate "shtt.mkd", blRepopRequired
End Sub
Private Sub mAddToVendorUpdateAsNeeded(ByRef ilVendorCodesToUpdate() As Integer)
    '8862
    Dim ilLoop As Integer
    Dim ilVendorLoop As Integer
    Dim ilUpdateCodeLoop As Integer
    Dim blFound As Boolean
    
On Error GoTo errbox
    If UBound(tmAutoUpdateVendors) > 0 Then
        For ilLoop = 0 To UBound(tmAirSpotInfo) - 1 Step 1
            'UD = undefined.  All blank were set to this
            If tmAirSpotInfo(ilLoop).sVendorSource <> "UD" And tmAirSpotInfo(ilLoop).sVendorSource <> "UI" Then
                'do we have an autoupdate vendor?  then let's see if this source matches
                For ilVendorLoop = 0 To UBound(tmAutoUpdateVendors) - 1 Step 1
                    If tmAirSpotInfo(ilLoop).sVendorSource = tmAutoUpdateVendors(ilVendorLoop).sSourceName Then
                        blFound = False
                        'have we already added it to updatecode array?
                        For ilUpdateCodeLoop = 0 To UBound(ilVendorCodesToUpdate) - 1
                            If ilVendorCodesToUpdate(ilUpdateCodeLoop) = tmAutoUpdateVendors(ilVendorLoop).iIdCode Then
                                blFound = True
                                Exit For
                            End If
                        Next ilUpdateCodeLoop
                        If Not blFound Then
                            ilUpdateCodeLoop = UBound(ilVendorCodesToUpdate)
                            ilVendorCodesToUpdate(ilUpdateCodeLoop) = tmAutoUpdateVendors(ilVendorLoop).iIdCode
                            ReDim Preserve ilVendorCodesToUpdate(ilUpdateCodeLoop + 1)
                        End If
                        Exit For
                    End If
                Next ilVendorLoop
            End If
        Next ilLoop
    End If
    Exit Sub
errbox:
    SetResults "warning!  Issue with Auto Vendor Delivery Status. Please see log", MESSAGERED
   ' myErrors.WriteError "mAddToVendorUpdateAsNeeded-" & Err.Description, False
    gLogMsg "mAddToVendorUpdateAsNeeded-" & Err.Description, cmPathForgLogMsg, False
End Sub


