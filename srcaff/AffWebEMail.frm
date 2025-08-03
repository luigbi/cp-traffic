VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmWebEMail 
   Caption         =   "Emails by Vehicle"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   Icon            =   "AffWebEMail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   10755
   Begin VB.Timer tmcLoadMessages 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9990
      Top             =   8070
   End
   Begin VB.CheckBox chkCombineEmails 
      Caption         =   "For any stations carrying more than one selected vehicle, eliminate duplicate emails"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   6855
   End
   Begin VB.CheckBox ckcCustSubject 
      Caption         =   "Use Custom Subject Line:"
      Height          =   435
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox txtCustSubject 
      Height          =   375
      Left            =   2760
      MaxLength       =   75
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.TextBox txtCCEmail 
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   3360
      Width           =   5655
   End
   Begin VB.TextBox txtActiveDate 
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CheckBox ckcCurrStations 
      Caption         =   "Include Stations Active On or After"
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   2790
      Width           =   3660
   End
   Begin VB.Frame frcSuppress 
      Caption         =   "Honor ""Suppress Overdue Notices"""
      Height          =   675
      Left            =   7920
      TabIndex        =   6
      Top             =   105
      Width           =   2745
      Begin VB.OptionButton rbcSuppress 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   315
         Width           =   735
      End
      Begin VB.OptionButton rbcSuppress 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Frame frcMessage 
      Caption         =   "Message"
      Height          =   1725
      Left            =   75
      TabIndex        =   9
      Top             =   825
      Width           =   10455
      Begin VB.CommandButton cmcSave 
         Caption         =   "Sa&ve"
         Height          =   375
         Left            =   9405
         TabIndex        =   31
         Top             =   1215
         Width           =   930
      End
      Begin VB.ListBox lbcVehiclesMsg 
         Height          =   1185
         ItemData        =   "AffWebEMail.frx":08CA
         Left            =   330
         List            =   "AffWebEMail.frx":08CC
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   330
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.TextBox edcMessage 
         Height          =   1275
         Left            =   345
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   330
         Width           =   8820
      End
      Begin VB.Image imcSpellCheck 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   30
         Picture         =   "AffWebEMail.frx":08CE
         ToolTipText     =   "Check Spelling"
         Top             =   1290
         Width           =   360
      End
   End
   Begin VB.Frame frcEMailType 
      Caption         =   "Message to be Sent"
      Height          =   675
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   7800
      Begin VB.OptionButton rbcEMailType 
         Caption         =   "Missed"
         Height          =   195
         Index           =   5
         Left            =   5760
         TabIndex        =   32
         Top             =   315
         Width           =   800
      End
      Begin VB.OptionButton rbcEMailType 
         Caption         =   "Welcome by Vehicle"
         Height          =   195
         Index           =   4
         Left            =   1365
         TabIndex        =   2
         Top             =   315
         Width           =   1770
      End
      Begin VB.OptionButton rbcEMailType 
         Caption         =   "Custom"
         Height          =   195
         Index           =   3
         Left            =   6720
         TabIndex        =   5
         Top             =   315
         Width           =   900
      End
      Begin VB.OptionButton rbcEMailType 
         Caption         =   "Overdue"
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   4
         Top             =   315
         Width           =   1305
      End
      Begin VB.OptionButton rbcEMailType 
         Caption         =   "Password"
         Height          =   195
         Index           =   1
         Left            =   3375
         TabIndex        =   3
         Top             =   315
         Width           =   1425
      End
      Begin VB.OptionButton rbcEMailType 
         Caption         =   "Welcome"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   315
         Width           =   1755
      End
   End
   Begin VB.ListBox lbcMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "AffWebEMail.frx":0F40
      Left            =   6585
      List            =   "AffWebEMail.frx":0F42
      TabIndex        =   29
      Top             =   5310
      Width           =   3945
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   720
      Top             =   7680
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   8610
      FormDesignWidth =   10755
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Send"
      Height          =   375
      Left            =   2460
      TabIndex        =   26
      Top             =   7905
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6420
      TabIndex        =   27
      Top             =   7905
      Width           =   1575
   End
   Begin VB.Frame frcTo 
      Caption         =   "Send Message To"
      Height          =   2925
      Left            =   75
      TabIndex        =   19
      Top             =   4920
      Width           =   6270
      Begin VB.ListBox lbcVehicles 
         Height          =   1815
         ItemData        =   "AffWebEMail.frx":0F44
         Left            =   135
         List            =   "AffWebEMail.frx":0F46
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   660
         Width           =   3855
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "All"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   2550
         Width           =   900
      End
      Begin VB.ListBox lbcStation 
         Height          =   1815
         ItemData        =   "AffWebEMail.frx":0F48
         Left            =   4215
         List            =   "AffWebEMail.frx":0F4A
         MultiSelect     =   2  'Extended
         TabIndex        =   24
         Top             =   660
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkAllStation 
         Caption         =   "All"
         Height          =   195
         Left            =   4230
         TabIndex        =   25
         Top             =   2550
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lacTitle1 
         Alignment       =   2  'Center
         Caption         =   "Vehicles"
         Height          =   285
         Left            =   90
         TabIndex        =   20
         Top             =   315
         Width           =   3885
      End
      Begin VB.Label lacTitle3 
         Alignment       =   2  'Center
         Caption         =   "Stations"
         Height          =   285
         Left            =   4170
         TabIndex        =   23
         Top             =   315
         Visible         =   0   'False
         Width           =   1740
      End
   End
   Begin VB.Label lblCCEmail 
      Caption         =   "Send Bcc E-mail To:"
      Height          =   435
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   120
      TabIndex        =   30
      Top             =   7845
      Width           =   5490
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   315
      Left            =   7200
      TabIndex        =   28
      Top             =   4920
      Width           =   1965
   End
End
Attribute VB_Name = "frmWebEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'******************************************************
'*  frmExport - Export ISCI
'*
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private smDate As String     'Export Date
Private imNumberDays As Integer
Private imVefArray() As Integer
Private lmArttCode() As Long
Private imVefCode As Integer
Private imAdfCode As Integer
Private smVefName As String
Private imAllClick As Integer
Private imAllStationClick As Integer
Private imExporting As Integer
Private imTerminate As Integer
'Private hmMsg As Integer
Private hmTo As Integer
Private hmToBody As Integer
Private smWelcome As String
Private smPassword As String
Private smOverdue As String
Private smSvWelcome As String
Private smSvPassword As String
Private smSvOverdue As String
Private imOMNoWeeks As Integer
Private smCustom As String
Private tmCPDat() As DAT
Private tmEMailRef() As EMAILREF
Private commrst As ADODB.Recordset
'10657 renamed
Private smToFileBody As String
Private smToFile As String
'10657
Private smMsgType As String
Private smCCEMail As String
Private smOMMinDate As String
Private imNoEmailAddr As Integer
'9926
Private smMissed As String
Private smSvMissed As String
Private Type WMVINFO
    iCmtCode As Integer
    iVefCode As Integer
    bChg As Boolean
    sComment As String * 1000
End Type
Private tmWMVInfo() As WMVINFO 'Private smCompliantBy As String 'Not Used with v7.0; it was: A=Advertiser; P=Pledge
Private imCurrentSelectedVehicle As Integer
Private bmInClick As Boolean
'10657
Private emEmailType As EMAILTYPE
Private Const WELCOME As Integer = 0
Private Const Password As Integer = 1
Private Const OVERDUE As Integer = 2
Private Const CUSTOM As Integer = 3
Private Const WELCOMEBYVEHICLE As Integer = 4
Private Const MISSED As Integer = 5
Private Const ErrorLog As String = "WebEmailLog.Txt"
Private Const MESSAGERED As Long = 255
Private Const MESSAGEBLACK As Long = 0
Private Enum EMAILTYPE
    WELCOMETYPE
    PASSWORDTYPE
    OVERDUETYPE
    CUSTOMTYPE
    WELCOMEBYVEHICLETYPE
    MISSEDTYPE
    NOTYPE
End Enum
Private Sub SetResults(sMsg As String, lFGC As Long)
    lbcMsg.AddItem sMsg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    lbcMsg.ForeColor = lFGC
    DoEvents
End Sub

Private Sub mFillVehicle()
    Dim iLoop As Integer
    lbcVehicles.Clear
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    chkAll.Value = 0
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        'If (tgVehicleInfo(iLoop).sOLAExport <> "Y") Then
            lbcVehicles.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            lbcVehicles.ItemData(lbcVehicles.NewIndex) = tgVehicleInfo(iLoop).iCode
        'End If
    Next iLoop
End Sub

Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    On Error GoTo ErrHand
    
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        iValue = True
        If lbcVehicles.ListCount > 1 Then
            lacTitle3.Visible = False
            chkAllStation.Visible = False
            lbcStation.Visible = False
            lbcStation.Clear
        Else
            lacTitle3.Visible = True
            chkAllStation.Visible = True
            lbcStation.Visible = True
        End If
    Else
        iValue = False
    End If
    If lbcVehicles.ListCount > 0 Then
        imAllClick = True
        lRg = CLng(lbcVehicles.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcVehicles.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllClick = False
    End If
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebExportSchdSpot - mFTPFiles: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Exit Sub

End Sub
Private Function mFTPFiles() As Boolean
    '10657 also added 'boolean' to function
    Dim slRegSection As String
    Dim slEmailType As String
    Dim slFileName As String
    Dim slFileNameBody As String
    Dim slErrorMessage As String
    On Error GoTo ErrHand

    'because smToFile has path included.
    slFileNameBody = "WebMsg_" & smMsgType & "Body.txt"
    slFileName = "WebMsg_" & smMsgType & ".txt"
    slEmailType = smMsgType
    'different message from name of file
    If emEmailType = MISSEDTYPE Then
        slEmailType = "Unresolved missed"
    End If
    SetResults "Sending email " & slEmailType & " body file to web site.", MESSAGEBLACK
    If Not gFTPFileToWebServer(smToFileBody, slFileNameBody) Then
        slErrorMessage = "FAILED to FTP " & slFileNameBody
        SetResults slErrorMessage, MESSAGERED
        gLogMsg "ERROR: " & slErrorMessage, ErrorLog, False
        mFTPFiles = False
        Exit Function
    End If
    SetResults "Sending email " & slEmailType & " file to web site.", MESSAGEBLACK
    If Not gFTPFileToWebServer(smToFile, slFileName) Then
        slErrorMessage = "FAILED to FTP " & slFileName
        SetResults slErrorMessage, MESSAGERED
        gLogMsg "ERROR: " & slErrorMessage, ErrorLog, False
        Screen.MousePointer = vbDefault
        mFTPFiles = False
        Exit Function
    End If
    SetResults "Telling web site to import network message file(s).", 0
    If Not gExecExtStoredProc("Nothing.txt", "NetworkEmail.exe", False, False) Then
        SetResults "FAIL: Unable to tell Web site to import Network email file(s)...", MESSAGERED
        gLogMsg "ERROR: Unable to tell Web site to import Network email file(s)...", ErrorLog, False
        Screen.MousePointer = vbDefault
        mFTPFiles = False
        Exit Function
    End If
    mFTPFiles = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebEmail - mFTPFiles: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Exit Function
    
End Function

Private Sub chkAllStation_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If imAllStationClick Then
        Exit Sub
    End If
    If chkAllStation.Value = vbChecked Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcStation.ListCount > 0 Then
        imAllStationClick = True
        lRg = CLng(lbcStation.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcStation.hwnd, LB_SELITEMRANGE, iValue, lRg)
        imAllStationClick = False
    End If
    

End Sub


Private Sub ckcCurrStations_Click()

    If ckcCurrStations.Value = vbUnchecked Then
        txtActiveDate.Visible = False
    Else
        txtActiveDate.Visible = True
    End If
        
    mFillStations

End Sub


Private Sub ckcCustSubject_Click()

    If ckcCustSubject.Value = vbChecked Then
        txtCustSubject.Visible = True
    Else
        txtCustSubject.Visible = False
    End If

End Sub

Private Sub cmcSave_Click()
    Dim ilCmtCode As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim blVehicleDeleteMsg As Boolean
    
    If sgUstWin(11) <> "I" Then
        Exit Sub
    End If
    blVehicleDeleteMsg = False
    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If Trim$(smSvWelcome) <> Trim$(smWelcome) Then
            If Trim$(smWelcome) <> "" Then
                ilCmtCode = rst!siteWMCmtCode
                ilRet = mSaveMessage(smWelcome, ilCmtCode, "W", 0)
                If (ilRet) And (rst!siteWMCmtCode <= 0) Then
                    SQLQuery = "Update Site Set "
                    SQLQuery = SQLQuery & "siteWMCmtCode = " & ilCmtCode
                    SQLQuery = SQLQuery & "Where siteCode = " & 1
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "frmWebEMail-cmcSave_Click"
                        Exit Sub
                    End If
                End If
            Else
                If Trim$(smSvWelcome) <> "" Then
                    MsgBox "Not Allowed to Remove the Welcome Message"
                End If
            End If
        End If
        If Trim$(smSvPassword) <> Trim$(smPassword) Then
            If Trim$(smPassword) <> "" Then
                ilCmtCode = rst!sitePMCmtCode
                ilRet = mSaveMessage(smPassword, ilCmtCode, "P", 0)
                If (ilRet) And (rst!sitePMCmtCode <= 0) Then
                    SQLQuery = "Update Site Set "
                    SQLQuery = SQLQuery & "sitePMCmtCode = " & ilCmtCode
                    SQLQuery = SQLQuery & "Where siteCode = " & 1
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "frmWebEMail-cmcSave_Click"
                        Exit Sub
                    End If
                End If
            Else
                If Trim$(smSvPassword) <> "" Then
                    MsgBox "Not Allowed to Remove the Password Message"
                End If
            End If
        End If
        If Trim$(smSvOverdue) <> Trim$(smOverdue) Then
            If Trim$(smOverdue) <> "" Then
                ilCmtCode = rst!siteOMCmtCode
                ilRet = mSaveMessage(smOverdue, ilCmtCode, "O", 0)
                If (ilRet) And (rst!siteOMCmtCode <= 0) Then
                    SQLQuery = "Update Site Set "
                    SQLQuery = SQLQuery & "siteOMCmtCode = " & ilCmtCode
                    SQLQuery = SQLQuery & "Where siteCode = " & 1
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "frmWebEMail-cmcSave_Click"
                        Exit Sub
                    End If
                End If
            Else
                If Trim$(smSvOverdue) <> "" Then
                    MsgBox "Not Allowed to Remove the Overdue Message"
                End If
            End If
        End If
        '9926 10028
        If Trim$(smSvMissed) <> Trim$(smMissed) Then
            If Trim$(smMissed) <> "" Then
                ilCmtCode = rst!siteUMCmtCode
                ilRet = mSaveMessage(smMissed, ilCmtCode, "M", 0)
                If (ilRet) And (rst!siteUMCmtCode <= 0) Then
                    SQLQuery = "Update Site Set "
                    SQLQuery = SQLQuery & "siteUMCmtCode = " & ilCmtCode
                    SQLQuery = SQLQuery & "Where siteCode = " & 1
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "frmWebEMail-cmcSave_Click"
                        Exit Sub
                    End If
                End If
            Else
                If Trim$(smSvMissed) <> "" Then
                    MsgBox "Not Allowed to Remove the Missed Message"
                End If
            End If
        End If
        
        smSvWelcome = Trim$(smWelcome)
        smSvPassword = Trim$(smPassword)
        smSvOverdue = Trim$(smOverdue)
        '9926
        smSvMissed = Trim$(smMissed)
        If rbcEMailType(WELCOMEBYVEHICLE).Value Then
            mRetainWMVInfo False
        End If
        For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
            If tmWMVInfo(ilVef).bChg Then
                If Trim$(tmWMVInfo(ilVef).sComment) <> "" Then
                    ilRet = mSaveMessage(tmWMVInfo(ilVef).sComment, tmWMVInfo(ilVef).iCmtCode, "V", tmWMVInfo(ilVef).iVefCode)
                    tmWMVInfo(ilVef).bChg = False
                    'If Not ilRet Then
                    '    Exit Function
                    'End If
                Else
                    If Not blVehicleDeleteMsg Then
                        MsgBox "Not Allowed to Remove the Vehicle Message"
                        blVehicleDeleteMsg = True
                    End If
                End If
            End If
        Next ilVef

        cmcSave.Enabled = False
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmfrmWebEMail-mSaveMessage"
End Sub

Private Sub cmdExport_Click()
    Dim iLoop As Integer
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim sMoDate As String
    Dim sNowDate As String
    Dim iIndex As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim slStr As String
    '10657
    Dim slErrorMessage As String
    
    On Error GoTo ErrHand
    
    If sgUstWin(11) <> "I" Then
        Exit Sub
    End If
    'Get all of the latest passwords and email addresses from the web
    gRemoteTestForNewEmail
    gRemoteTestForNewWebPW
    smCCEMail = Trim(txtCCEMail.Text)
    iRet = gTestForMultipleEmail(smCCEMail, "BCC")
    If iRet = False Then
        Screen.MousePointer = vbDefault
        gMsgBox sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Send Bcc Email To Address Before Continuing", vbExclamation
        gLogMsg sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Send Bcc Email To Address Before Continuing", ErrorLog, False
        txtCCEMail.SetFocus
        Exit Sub
    End If
    If ckcCurrStations.Value = vbChecked And Not rbcEMailType(OVERDUE).Value Then
        ilRet = mValidDate()
        If ilRet = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    lbcMsg.Clear
    lbcMsg.ForeColor = RGB(0, 0, 0)
    If lbcVehicles.ListIndex < 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    '10657 do this test just once.  welcome, overdue, etc.
    slErrorMessage = ""
    If Not mSetEmailTypeAndFileNames() Then
        If emEmailType = WELCOMEBYVEHICLETYPE Then
            slErrorMessage = "General welcome message must be defined as not all vehicles have a welcome message."
        Else
            slErrorMessage = "Please specify message."
        End If
    End If
    If emEmailType = NOTYPE Then
        slErrorMessage = "Please specify type of e-mail to be sent."
    ElseIf emEmailType = CUSTOMTYPE And ckcCustSubject.Value = vbChecked And Len(Trim(txtCustSubject.Text)) = 0 Then
        slErrorMessage = "Please Specify the Custom Subject."
    End If
    If Len(slErrorMessage) > 0 Then
        Screen.MousePointer = vbDefault
        Beep
        gMsgBox slErrorMessage, vbCritical
        edcMessage.SetFocus
        Exit Sub
    End If
    '9926
'    If (rbcEMailType(WELCOME).Value = False) And (rbcEMailType(Password).Value = False) And (rbcEMailType(OVERDUE).Value = False) And (rbcEMailType(CUSTOM).Value = False) And (rbcEMailType(WELCOMEBYVEHICLE).Value = False) And (rbcEMailType(MISSED).Value = False) Then
'        Screen.MousePointer = vbDefault
'        Beep
'        gMsgBox "Please Specify type of E-Mail to be Sent.", vbCritical
'        Exit Sub
'    End If
'    If rbcEMailType(WELCOME).Value Then
'        If Len(Trim(smWelcome)) = 0 Then
'            Screen.MousePointer = vbDefault
'            Beep
'            gMsgBox "Please Specify Message.", vbCritical
'            edcMessage.SetFocus
'            Exit Sub
'        End If
'    ElseIf rbcEMailType(Password).Value Then
'        If Len(Trim(smPassword)) = 0 Then
'            Screen.MousePointer = vbDefault
'            Beep
'            gMsgBox "Please Specify Message.", vbCritical
'            edcMessage.SetFocus
'            Exit Sub
'        End If
'    ElseIf rbcEMailType(OVERDUE).Value Then
'        If Len(Trim(smOverdue)) = 0 Then
'            Screen.MousePointer = vbDefault
'            Beep
'            gMsgBox "Please Specify Message.", vbCritical
'            edcMessage.SetFocus
'            Exit Sub
'        End If
'    ElseIf rbcEMailType(CUSTOM).Value Then
'        If Len(Trim(smCustom)) = 0 Then
'            Screen.MousePointer = vbDefault
'            Beep
'            gMsgBox "Please Specify Message.", vbCritical
'            edcMessage.SetFocus
'            Exit Sub
'        End If
'
'        If ckcCustSubject.Value = vbChecked Then
'            If Len(Trim(txtCustSubject.Text)) = 0 Then
'                Screen.MousePointer = vbDefault
'                gMsgBox "Please Specify the Custom Subject.", vbCritical
'                Exit Sub
'            End If
'        End If
'    ElseIf rbcEMailType(WELCOMEBYVEHICLE).Value Then
'        chkCombineEmails.Value = vbUnchecked
'        mRetainWMVInfo False
'        If Len(Trim(smWelcome)) = 0 Then
'            For iLoop = 0 To lbcVehicles.ListCount - 1
'                If lbcVehicles.Selected(iLoop) Then
'                    'Get hmTo handle
'                    imVefCode = lbcVehicles.ItemData(iLoop)
'                    For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
'                        If tmWMVInfo(ilVef).iVefCode = imVefCode Then
'                            If Trim$(tmWMVInfo(ilVef).sComment) = "" Then
'                                Screen.MousePointer = vbDefault
'                                Beep
'                                '10657 misspelling
'                                gMsgBox "General welcome message must be defined as not all vehicles have a welcome message.", vbCritical
'                                Exit Sub
'                            End If
'                        End If
'                    Next ilVef
'                End If
'            Next iLoop
'        End If
'    '9926
'    ElseIf rbcEMailType(MISSED).Value Then
'        If Len(Trim(smMissed)) = 0 Then
'            Screen.MousePointer = vbDefault
'            Beep
'            gMsgBox "Please Specify Message.", vbCritical
'            edcMessage.SetFocus
'            Exit Sub
'        End If
'    End If
    Screen.MousePointer = vbHourglass
    
'    If Not mOpenMsgFile(sMsgFileName) Then
'        Screen.MousePointer = vbDefault
'        cmdCancel.SetFocus
'        Exit Sub
'    End If
    imExporting = True
    On Error GoTo 0
    ReDim tmEMailRef(0 To 0) As EMAILREF
    '10657
    If igDemoMode Then
        SetResults "In Demo Mode.  NO EMAILS WILL BE SENT", 0
    End If
    SetResults "Gathering Vehicle/Stations Information.", 0
    lacResult.Caption = ""
    For iLoop = 0 To lbcVehicles.ListCount - 1
        If lbcVehicles.Selected(iLoop) Then
            'Get hmTo handle
            imVefCode = lbcVehicles.ItemData(iLoop)
            smVefName = Trim$(lbcVehicles.List(iLoop))
            If sgShowByVehType = "Y" Then
                smVefName = Mid$(smVefName, 3)
            End If
            Screen.MousePointer = vbHourglass
            'ReDim imVefArray(0 To 1) As Integer
            'imVefArray(0) = imVefCode
            'ilVef = gBinarySearchVef(CLng(imVefCode))
            'If ilVef <> -1 Then
            '    If tgVehicleInfo(ilVef).sVehType = "L" Then
            '        ilVff = gBinarySearchVff(imVefCode)
            '        If ilVff <> -1 Then
            '            slStr = Trim$(tgVffInfo(ilVff).sWebName)
            '            If slStr <> "" Then
            '                smVefName = slStr
            '                If sgShowByVehType = "Y" Then
            '                    smVefName = Mid$(smVefName, 3)
            '                End If
            '            End If
            '        End If
            '        For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) Step 1
            '            If tgVehicleInfo(ilVef).iVefCode = imVefCode Then
            '                For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
            '                    If tgVehicleInfo(ilVef).iCode = tgVffInfo(ilVff).iVefCode Then
            '                        If tgVffInfo(ilVff).sMergeWeb <> "S" Then
            '                            imVefArray(UBound(imVefArray)) = tgVehicleInfo(ilVef).iCode
            '                            ReDim Preserve imVefArray(0 To (UBound(imVefArray) + 1)) As Integer
            '                            Exit For
            '                        End If
            '                    End If
            '                Next ilVff
            '            End If
            '        Next ilVef
            '    End If
            'End If
            mBuildVehicleArray
            'Gather info
            ReDim lmArttCode(0 To 0) As Long
            For ilVef = 0 To UBound(imVefArray) - 1 Step 1
                imVefCode = imVefArray(ilVef)   'Leave vehicle name as set above (smVefName)
                iRet = mGatherInfo()
                If (iRet = False) Then
                    SetResults "FAILED: Gather Vehicle/Station Info", MESSAGERED
                    gLogMsg "ERROR: FAILED to Gather Vehicle/Station Info", ErrorLog, False
                    'Print #hmMsg, "** Terminated **"
                    'Close #hmMsg
                    imExporting = False
                    Screen.MousePointer = vbDefault
                    cmdCancel.SetFocus
                    Exit Sub
                End If
                If imTerminate Then
                    SetResults "User Terminate: Gather Vehicle/Station Info", MESSAGERED
                    gLogMsg "User Terminate: Gather Vehicle/Station Info", ErrorLog, False
                    'Print #hmMsg, "** User Terminated **"
                    'Close #hmMsg
                    imExporting = False
                    Screen.MousePointer = vbDefault
                    cmdCancel.SetFocus
                    Exit Sub
                End If
            Next ilVef
            If emEmailType = WELCOMEBYVEHICLETYPE Then
                SetResults "Sending E-Mailing Information for " & smVefName, 0
                DoEvents
                '10657
                If UBound(tmEMailRef) > LBound(tmEMailRef) Then
                    If Not mCreateAndFTPFiles() Then
                        imExporting = False
                        Screen.MousePointer = vbDefault
                        cmdCancel.SetFocus
                        Exit Sub
                    Else
                        ReDim tmEMailRef(0 To 0) As EMAILREF
                    End If
                End If
            '9926  not based on vehicles.  Only needs to run once.
            ElseIf emEmailType = MISSEDTYPE Then
                Exit For
            End If
       End If
    Next iLoop
    'Sort Cross Reference
    If UBound(tmEMailRef) - 1 > 0 Then
        ArraySortTyp fnAV(tmEMailRef(), 0), UBound(tmEMailRef), 0, LenB(tmEMailRef(0)), 0, LenB(tmEMailRef(0).sKey), 0
    End If
    If Not emEmailType = WELCOMEBYVEHICLETYPE Then
        SetResults "E-Mailing Vehicle/Stations Information.", 0
        '10657
        If UBound(tmEMailRef) > LBound(tmEMailRef) Then
            If Not mCreateAndFTPFiles() Then
                imExporting = False
                Screen.MousePointer = vbDefault
                cmdCancel.SetFocus
                Exit Sub
            End If
        'these files may not exist-not an issue.
        ElseIf emEmailType = OVERDUETYPE Or emEmailType = MISSEDTYPE Then
            SetResults "No " & smMsgType & " messages to send.", 0
        End If
    End If
    imExporting = False
    'Print #hmMsg, "** Completed Web E-Mail: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    gLogMsg "** Completed Web E-Mail **", ErrorLog, False
    SetResults "E-Mail Completed Successfully.", RGB(0, 155, 0)
    '10657
    If igDemoMode Then
        SetResults "In Demo Mode.  NO EMAILS WERE SENT", RGB(0, 155, 0)
    End If
    lbcMsg.ListIndex = -1   ' Finish with nothing selected
    'Close #hmMsg '9926 missed never comes in here
    '10657 this does nothing commented out
'    If chkCombineEmails.Value And (Not rbcEMailType(WELCOMEBYVEHICLE).Value) Then
'        iRet = mPostProcessFile()
'    End If
    lacResult.Caption = "Results: " & sMsgFileName
    cmdExport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    If imNoEmailAddr Then
        gMsgBox "Some station(s) did not have an email address defined.  No email can be sent to those stations.  Please refer to your Messages folder for a listing.  See file WebNoEmailAddress.Txt for results.", vbExclamation
    End If
    
    Exit Sub
cmdExportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebEmail - cmdExport click: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmWebEMail
End Sub


Private Sub edcMessage_Change()
'    If rbcEMailType(WELCOME).Value Then
'        smWelcome = edcMessage.Text
'    ElseIf rbcEMailType(PASSWORD).Value Then
'        smPassword = edcMessage.Text
'    ElseIf rbcEMailType(OVERDUE).Value Then
'        smOverdue = edcMessage.Text
'    ElseIf rbcEMailType(CUSTOM).Value Then
'        smCustom = edcMessage.Text
'    End If
End Sub

Private Sub edcMessage_LostFocus()
    If rbcEMailType(WELCOME).Value Then
        If smWelcome <> edcMessage.Text Then
            smWelcome = edcMessage.Text
            cmcSave.Enabled = True
        End If
    ElseIf rbcEMailType(Password).Value Then
        If smPassword <> edcMessage.Text Then
            smPassword = edcMessage.Text
            cmcSave.Enabled = True
        End If
    ElseIf rbcEMailType(OVERDUE).Value Then
        If smOverdue <> edcMessage.Text Then
            smOverdue = edcMessage.Text
            cmcSave.Enabled = True
        End If
    ElseIf rbcEMailType(CUSTOM).Value Then
        smCustom = edcMessage.Text
    ElseIf rbcEMailType(WELCOMEBYVEHICLE).Value Then
        mRetainWMVInfo True
    '9926
    ElseIf rbcEMailType(MISSED).Value Then
        smMissed = edcMessage.Text
         cmcSave.Enabled = True
    End If
    If sgUstWin(11) <> "I" Then
        cmcSave.Enabled = False
    End If
End Sub

Private Sub Form_Initialize()
    If Not Me.WindowState = vbMaximized Then
        'Me.Width = Screen.Width / 1.05
        Me.Width = Screen.Width / 1.3
        'Me.Height = Screen.Height / 1.7
        Me.Height = Screen.Height / 1.5
        Me.Top = (Screen.Height - Me.Height) / 2
        Me.Left = (Screen.Width - Me.Width) / 2
        gCenterForm frmWebEMail
    End If
    gSetFonts frmWebEMail
End Sub

Private Sub Form_Load()
    Dim iRet As Integer
    
    Screen.MousePointer = vbHourglass
    bgEMailVisible = True
    frcSuppress.Visible = False
    rbcSuppress(0).Value = True
    imAllClick = False
    imAllStationClick = False
    imTerminate = False
    imExporting = False
    iRet = mInit()
    lbcStation.Clear
    mFillVehicle
    chkAll.Value = vbChecked
    ckcCurrStations.Value = vbChecked
    txtCustSubject.Visible = False
    ckcCustSubject.Visible = False
    txtActiveDate.Text = Format(gNow(), "mm/dd/yy")
    tmcLoadMessages.Enabled = True
    cmcSave.Enabled = False
    If sgUstWin(11) <> "I" Then
        cmdExport.Enabled = False
    End If
    
    If gIsUsingNovelty Then
        rbcEMailType(Password).Visible = False
        rbcEMailType(OVERDUE).Left = rbcEMailType(OVERDUE).Left - 1400
        rbcEMailType(MISSED).Left = rbcEMailType(MISSED).Left - 1400
        rbcEMailType(CUSTOM).Left = rbcEMailType(CUSTOM).Left - 1400
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    bgEMailVisible = False
    Erase lmArttCode
    Erase imVefArray
    Erase tmCPDat
    Erase tmEMailRef
    Erase tmWMVInfo
    commrst.Close
    Set frmWebEMail = Nothing
End Sub


Private Sub imcSpellCheck_Click()
    gSpellCheckUsingMSWord edcMessage
End Sub

Private Sub lbcStation_Click()
    
    lbcMsg.Clear
    If imAllStationClick Then
        Exit Sub
    End If
    If (cmdExport.Enabled = False) And (sgUstWin(11) = "I") Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If chkAllStation.Value = vbChecked Then
        imAllStationClick = True
        chkAllStation.Value = vbUnchecked
        imAllStationClick = False
    End If
End Sub

Private Sub lbcVehicles_Click()
    Dim iLoop As Integer
    Dim iCount As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
'    If ckcCurrStations.Value = vbChecked And Not rbcEMailType(OVERDUE).Value Then
'        ilRet = mValidDate()
'        If ilRet = False Then
'            Exit Sub
'        End If
'    End If
    
    lbcMsg.Clear
    lbcStation.Clear
    If (cmdExport.Enabled = False) And (sgUstWin(11) = "I") Then
        cmdExport.Enabled = True
        cmdCancel.Caption = "&Cancel"
    End If
    If chkAllStation.Value = vbChecked Then
        chkAllStation.Value = vbUnchecked
    End If
    If imAllClick Then
        Exit Sub
    End If
    If chkAll.Value = vbChecked Then
        imAllClick = True
        chkAll.Value = vbUnchecked
        imAllClick = False
    End If
    For iLoop = 0 To lbcVehicles.ListCount - 1 Step 1
        If lbcVehicles.Selected(iLoop) Then
            imVefCode = lbcVehicles.ItemData(iLoop)
            iCount = iCount + 1
            If iCount > 1 Then
                Exit For
            End If
        End If
    Next iLoop
    If iCount = 1 Then
        lacTitle3.Visible = True
        chkAllStation.Visible = True
        lbcStation.Visible = True
        mFillStations
    Else
        lacTitle3.Visible = False
        chkAllStation.Visible = False
        lbcStation.Visible = False
    End If
    
    If lbcVehicles.SelCount < 2 Then
        chkCombineEmails.Visible = False
        chkCombineEmails.Value = vbUnchecked
    Else
        chkCombineEmails.Visible = True
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmWebEmail - lbcVehicles click: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    Exit Sub
    
End Sub

Private Sub mFillStations()

    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim llRow As Long
    
    On Error GoTo ErrHand
    
    lbcStation.Clear
    If ckcCurrStations.Value = vbChecked Then
        If Not mValidDate Then
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    mBuildVehicleArray
    For ilVef = 0 To UBound(imVefArray) - 1 Step 1
        imVefCode = imVefArray(ilVef)
        SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
        SQLQuery = SQLQuery + " FROM shtt, att"
        SQLQuery = SQLQuery + " WHERE (attVefCode = " & imVefCode
    '    SQLQuery = SQLQuery + " AND attExportType = 2 "
        SQLQuery = SQLQuery + " AND shttCode = attShfCode)"
        'D.S. 2/28/05 If checked only allow active agreements
        If ckcCurrStations.Value = vbChecked And Not rbcEMailType(OVERDUE).Value Then
            SQLQuery = SQLQuery + " AND (attOffAir >= '" & Format$(txtActiveDate.Text, sgSQLDateForm) & "') And (attDropDate >= '" & Format$(txtActiveDate.Text, sgSQLDateForm) & "') "
        End If
        SQLQuery = SQLQuery + " ORDER BY shttCallLetters"
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            llRow = SendMessageByString(lbcStation.hwnd, LB_FINDSTRING, -1, Trim$(rst!shttCallLetters))
            If llRow < 0 Then
                lbcStation.AddItem Trim$(rst!shttCallLetters)
                lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
            End If
            rst.MoveNext
        Wend
    Next ilVef
    chkAllStation.Value = vbChecked
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebEMail-mFillStation"
End Sub

Private Function mCreateFiles() As Integer
    Dim iRet As Integer
    Dim sDateTime As String
    Dim ilIndex As Integer
    Dim sStrToPrint As String
    Dim slStr As String
    Dim slNewVehList As String
    Dim slVehArray() As String
    Dim ilLoopDupes As Integer
    Dim slDates As String
    Dim blNewVehicle As Boolean
    
    mCreateFiles = False
    sStrToPrint = ""
    If Not mDeletePreviousWebMsgFiles(smToFile, smToFileBody) Then
        Close hmTo
        Close hmToBody
        Exit Function
    End If
    'open and write the header and basic
    iRet = gFileOpen(smToFile, "Output", hmTo)
    If iRet <> 0 Then
        Close hmTo
        '9960
        Close hmToBody
        gMsgBox "Open File " & smToFile & " error#" & Str$(Err.Number), vbOKOnly
        Exit Function
    End If
    iRet = gFileOpen(smToFileBody, "Output", hmToBody)
    If iRet <> 0 Then
        Close hmTo
        Close hmToBody
        gMsgBox "Open File " & smToFileBody & " error#" & Str$(Err.Number), vbOKOnly
        Exit Function
    End If
    mBasicAndHeaderForFile
    On Error GoTo ErrHand
    ReDim slVehArray(0 To 0)
    blNewVehicle = True
    For ilIndex = 0 To UBound(tmEMailRef) - 1 Step 1
        DoEvents
        If imTerminate Then
            Close hmTo
            Close hmToBody
            Exit Function
        End If
        If Len(Trim$(tmEMailRef(ilIndex).sWebEMail)) <> 0 Then
            Select Case emEmailType
                Case WELCOMEBYVEHICLETYPE, MISSEDTYPE
                    sStrToPrint = mLineForFile(ilIndex, True)
                    Print #hmTo, sStrToPrint
                Case OVERDUETYPE
                    If chkCombineEmails.Value Then
                        'vehicles will contain all vehicles for a station.  Each vehicle will be followed by all its overdue dates.  Each vehicle is separated by 2 carriage returns.
                        'the 'dates' field will be empty.
                        If blNewVehicle Then
                            blNewVehicle = False
                            slNewVehList = slNewVehList & " " & Trim(tmEMailRef(ilIndex).sVehName)
                        End If
                        'Write the line with station and emails the last time station found
                        If StrComp(tmEMailRef(ilIndex).sCallLetters, tmEMailRef(ilIndex + 1).sCallLetters, vbTextCompare) = 0 Then
                            slDates = slDates & ", " & Format(tmEMailRef(ilIndex).lDate, "mm/dd/yyyy")
                            ' is it last of this vehicle?  Format the line with carriage returns
                            If tmEMailRef(ilIndex).iVefCode <> tmEMailRef(ilIndex + 1).iVefCode Then
                                'here's the additional space for dates on first line only.  It's now always there.
                                slNewVehList = slNewVehList & " " & slDates & vbCrLf & vbCrLf
                                slDates = ""
                                blNewVehicle = True
                            End If
                        Else
                            'get the last for this station
                            slDates = slDates & ", " & Format(tmEMailRef(ilIndex).lDate, "mm/dd/yyyy")
                            slNewVehList = slNewVehList & " " & slDates
                            sStrToPrint = mLineForFile(ilIndex, False, True)
                            'Where Dates should be is a single space
                            sStrToPrint = """" & slNewVehList & """" & "," & sStrToPrint & """" & " " & """"
                            Print #hmTo, sStrToPrint
                            slNewVehList = ""
                            slDates = ""
                            blNewVehicle = True
                        End If
                    Else
                        If ilIndex > LBound(tmEMailRef) Then
                            If (StrComp(tmEMailRef(ilIndex - 1).sVehName, tmEMailRef(ilIndex).sVehName, vbTextCompare) <> 0) Or (StrComp(tmEMailRef(ilIndex - 1).sCallLetters, tmEMailRef(ilIndex).sCallLetters, vbTextCompare) <> 0) Then
                                If sStrToPrint <> "" Then
                                    sStrToPrint = sStrToPrint & """"
                                    Print #hmTo, sStrToPrint
                                End If
                                sStrToPrint = mLineForFile(ilIndex, True)
                            Else
                                sStrToPrint = sStrToPrint & ", " & Format(tmEMailRef(ilIndex).lDate, "mm/dd/yyyy")
                            End If
                            ' print if last one
                            If ilIndex + 1 = UBound(tmEMailRef) Then
                                sStrToPrint = sStrToPrint & """"
                                Print #hmTo, sStrToPrint
                            End If
                        'first time
                        Else
                            sStrToPrint = mLineForFile(ilIndex, True)
                        End If
                    End If
                'Custom, Password and welcome
                Case Else
                    If chkCombineEmails.Value Then
                        If StrComp(tmEMailRef(ilIndex).sCallLetters, tmEMailRef(ilIndex + 1).sCallLetters, vbTextCompare) = 0 Then
                            slVehArray(UBound(slVehArray)) = Trim(tmEMailRef(ilIndex).sVehName)
                            ReDim Preserve slVehArray(UBound(slVehArray) + 1)
                        Else
                            For ilLoopDupes = 0 To UBound(slVehArray) - 1
                                If ilLoopDupes = 0 Then
                                    slNewVehList = Trim$(slVehArray(ilLoopDupes))
                                Else
                                    If ilLoopDupes < UBound(slVehArray) Then
                                        slNewVehList = slNewVehList & ", " & Trim(slVehArray(ilLoopDupes))
                                    Else
                                        slNewVehList = slNewVehList & Trim(slVehArray(ilLoopDupes))
                                    End If
                                End If
                            Next ilLoopDupes
                            If ilLoopDupes = 0 Then
                                slNewVehList = Trim(tmEMailRef(ilIndex).sVehName)
                            Else
                                slNewVehList = slNewVehList & ", " & Trim(tmEMailRef(ilIndex).sVehName)
                            End If
                            sStrToPrint = """" & slNewVehList & """" & "," & mLineForFile(ilIndex, False)
                            Print #hmTo, sStrToPrint
                            slNewVehList = ""
                            ReDim slVehArray(0 To 0)
                        End If
                    Else
                        sStrToPrint = mLineForFile(ilIndex, True)
                        Print #hmTo, sStrToPrint
                    End If
            End Select
        End If
    Next
    Close #hmTo
    Close hmToBody
    mCreateFiles = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmCreateFiles-mCreateFiles"
    mCreateFiles = False
    Exit Function
End Function
Private Function mGatherInfo() As Integer

    Dim sCallLetters As String
    Dim sWebEMail As String
    Dim sWebPW As String
    Dim iShfCode As Integer
    Dim ilOkStation As Integer
    Dim iLoop As Integer
    Dim iFound As Integer
    Dim iIndex As Integer
    Dim iUpper As Integer
    Dim sDate As String
    Dim lDate As Long
    Dim ilRet As Integer
    Dim llIdx As Long
    Dim rst_Email As ADODB.Recordset
    Dim ilLen As Integer
    Dim blArttFd As Boolean
    Dim llArtt As Long
    '9926
    Dim slstations() As String
    Dim ilCounter As Integer
    Dim slStrippedStationName As String
    '9960
    Dim slTempEmail As String
    
    On Error GoTo ErrHand
    
    imNoEmailAddr = False
    llIdx = 0
    mGatherInfo = False
    If ckcCurrStations.Value = vbChecked Then
        If Not mValidDate() Then
            Exit Function
        End If
    End If
    
    If (rbcEMailType(WELCOME).Value) Or (rbcEMailType(Password).Value) Or (rbcEMailType(CUSTOM).Value) Or (rbcEMailType(WELCOMEBYVEHICLE).Value) Then
        DoEvents
        'Find all of the stations that have agreements that are active for the given vehicle
        SQLQuery = "SELECT shttCode, shttCallLetters, ShttWebPW"
        SQLQuery = SQLQuery + " FROM shtt, att"
        SQLQuery = SQLQuery + " WHERE (ShttCode = attShfCode)"
        If ckcCurrStations.Value = vbChecked Then
            SQLQuery = SQLQuery + " AND (attOffAir >= '" & Format$(txtActiveDate.Text, sgSQLDateForm) & "') And (attDropDate >= '" & Format$(txtActiveDate.Text, sgSQLDateForm) & "') "
        End If
        SQLQuery = SQLQuery + " AND (attvefCode = " & imVefCode & ")"
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            DoEvents
            If imTerminate Then
                Exit Function
            End If
            sCallLetters = Trim$(rst!shttCallLetters)
            iShfCode = rst!shttCode
            If lbcStation.ListCount > 0 Then
                ilOkStation = False
                For iLoop = 0 To lbcStation.ListCount - 1 Step 1
                    If lbcStation.Selected(iLoop) Then
                        If lbcStation.ItemData(iLoop) = iShfCode Then
                            ilOkStation = True
                            Exit For
                        End If
                    End If
                Next iLoop
            Else
                ilOkStation = True
            End If
            If ilOkStation Then
                'Get all of the addresses that are checked to recieve emails
                SQLQuery = "SELECT arttCode, arttEmail"
                SQLQuery = SQLQuery + " FROM artt"
                SQLQuery = SQLQuery + " WHERE (arttShttCode = " & rst!shttCode
                SQLQuery = SQLQuery + " AND arttWebEMail = 'Y'" & ")"
                Set rst_Email = gSQLSelectCall(SQLQuery)
                    
                sWebEMail = ""
                While Not rst_Email.EOF
                    DoEvents
                    If imTerminate Then
                        Exit Function
                    End If
                    'Remove duplicate references
                    blArttFd = False
                    For llArtt = 0 To UBound(lmArttCode) - 1 Step 1
                        If lmArttCode(llArtt) = rst_Email!arttCode Then
                            blArttFd = True
                            Exit For
                        End If
                    Next llArtt
                    '9960 stop blank email addresses which destroy the previous email address
'                    If Not blArttFd Then
'                        sWebEMail = sWebEMail & gFixQuote(Trim$(rst_Email!arttEmail))
'                        sWebEMail = sWebEMail & ","
'                        lmArttCode(UBound(lmArttCode)) = rst_Email!arttCode
'                        ReDim Preserve lmArttCode(0 To UBound(lmArttCode) + 1) As Long
'                    End If
                    If Not blArttFd Then
                        slTempEmail = gFixQuote(Trim$(rst_Email!arttEmail))
                        If Len(slTempEmail) > 0 Then
                            sWebEMail = sWebEMail & slTempEmail & ","
                            lmArttCode(UBound(lmArttCode)) = rst_Email!arttCode
                            ReDim Preserve lmArttCode(0 To UBound(lmArttCode) + 1) As Long
                        End If
                    End If
                    rst_Email.MoveNext
                Wend
        
                ilLen = Len(sWebEMail)
                If ilLen > 0 Then
                    sWebEMail = Left(sWebEMail, ilLen - 1)
                End If
                
                sWebPW = Trim$(rst!shttWebPW)
                
                'If Len(sWebEMail) = 0 Then
                '    sWebEMail = Trim$(rst!shttWebEmail)
                '    sWebPW = Trim$(rst!ShttWebPW)
                'End If

                ilRet = gTestForMultipleEmail(sWebEMail, "RegEmail")
                If ilRet = False Then
                    gMsgBox sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Email Address for Station: " & sCallLetters & " running on: " & smVefName & " Before Continuing", vbExclamation
                    gLogMsg sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Email Address for Station: " & sCallLetters & " running on: " & smVefName & " Before Continuing", ErrorLog, False
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
                
                iFound = False
                For iIndex = 0 To UBound(tmEMailRef) - 1 Step 1
                    If (StrComp(Trim$(tmEMailRef(iIndex).sVehName), smVefName, vbTextCompare) = 0) And (StrComp(Trim$(tmEMailRef(iIndex).sCallLetters), sCallLetters, vbTextCompare) = 0) Then
                        iFound = True
                        Exit For
                    End If
                Next iIndex
                If Not iFound Then
                    iUpper = UBound(tmEMailRef)
                    tmEMailRef(iUpper).sVehName = smVefName
                    tmEMailRef(iUpper).sCallLetters = sCallLetters
                    If Len(sWebEMail) = 0 Then
                        If Not imNoEmailAddr Then
                            gLogMsg "ERROR: " & sCallLetters & " Running on " & smVefName & " has no email address defined.", "WebNoEmailAddress.Txt", True
                        Else
                            gLogMsg "ERROR: " & sCallLetters & " Running on " & smVefName & " has no email address defined.", "WebNoEmailAddress.Txt", False
                        End If
                        imNoEmailAddr = True
                    End If
                    'D.S. 01-15-09
                    If Len(sWebEMail) > 5 Then
                        tmEMailRef(iUpper).sWebEMail = sWebEMail
                        tmEMailRef(iUpper).sWebPW = sWebPW
                        tmEMailRef(iUpper).iVefCode = imVefCode
                        tmEMailRef(iUpper).iShfCode = iShfCode
                        tmEMailRef(iUpper).lDate = 0
                        'tmEMailRef(iUpper).sKey = tmEMailRef(iUpper).sVehName & "|" & tmEMailRef(iUpper).sCallLetters
                        'D.S. 5/11/15 Flipped the .sKey from veh, callletters to callletters, veh when combine emails are checked
                        If chkCombineEmails.Value And (Not rbcEMailType(WELCOMEBYVEHICLE).Value) Then
                            tmEMailRef(iUpper).sKey = tmEMailRef(iUpper).sCallLetters & "|" & tmEMailRef(iUpper).sVehName
                        Else
                            tmEMailRef(iUpper).sKey = tmEMailRef(iUpper).sVehName & "|" & tmEMailRef(iUpper).sCallLetters
                        End If
                        ReDim Preserve tmEMailRef(0 To iUpper + 1) As EMAILREF
                    End If
                End If
            End If
            rst.MoveNext
        Wend
    ElseIf (rbcEMailType(OVERDUE).Value) Then
        sDate = Format$(gNow(), "m/d/yy")
        sDate = DateAdd("d", -1, sDate)
        
        Do While Weekday(sDate) <> vbSunday
            sDate = DateAdd("d", -1, sDate)
        Loop
        
        sDate = Format$(DateAdd("ww", -imOMNoWeeks, sDate), "mm/dd/yy")
        'SQLQuery = "SELECT shttCode, shttCallLetters, shttMarket, cpttStartDate, attWebEMail, attWebPW, attSuppressNotice"
        'SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        'SQLQuery = "SELECT shttCode, shttCallLetters, cpttStartDate, attWebEMail, attWebPW, attSuppressNotice, mktName"
        'SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, cptt, att"
        SQLQuery = "SELECT shttCode, shttCallLetters, ShttWebPW, cpttStartDate, attWebEMail, attWebPW, attSuppressNotice"
        SQLQuery = SQLQuery + " FROM shtt, cptt, att"
        SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
        SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
        SQLQuery = SQLQuery + " And attExportType = 1"
        SQLQuery = SQLQuery + " AND cpttStatus = 0"
        SQLQuery = SQLQuery + " AND cpttVefCode = " & imVefCode
        SQLQuery = SQLQuery & " AND cpttStartDate >= '" & Format$(smOMMinDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND cpttStartDate <= '" & Format$(sDate, sgSQLDateForm) & "')"
        Set rst = gSQLSelectCall(SQLQuery)
        
        '9/13/11: Doug-Remove this call as it is remade below
        'SQLQuery = "SELECT emtEmail"
        'SQLQuery = SQLQuery + " FROM emt"
        'SQLQuery = SQLQuery + " WHERE (emtShttCode = " & rst!shttCode & ")"
        'Set rst_Email = gSQLSelectCall(SQLQuery)
            
        While Not rst.EOF
            DoEvents
            If imTerminate Then
                Exit Function
            End If
            sCallLetters = Trim$(rst!shttCallLetters)
            
            'debug
'            If sCallLetters = "KATZ-AM" Then
'                ilRet = ilRet
'            End If

            
            iShfCode = rst!shttCode
            If lbcStation.ListCount > 0 Then
                ilOkStation = False
                For iLoop = 0 To lbcStation.ListCount - 1 Step 1
                    If lbcStation.Selected(iLoop) Then
                        If lbcStation.ItemData(iLoop) = iShfCode Then
                            ilOkStation = True
                            Exit For
                        End If
                    End If
                Next iLoop
            Else
                ilOkStation = True
            End If
            If ilOkStation And rbcSuppress(0).Value Then
                If rst!attSuppressNotice = "Y" Then
                    ilOkStation = False
                End If
            End If
            
            If ilOkStation Then
            
                '9/13/11: Doug- Replaced EMT with ARTT
                'SQLQuery = "SELECT emtEmail"
                'SQLQuery = SQLQuery + " FROM emt"
                'SQLQuery = SQLQuery + " WHERE (emtShttCode = " & rst!shttCode & ")"
                'Set rst_Email = gSQLSelectCall(SQLQuery)
                '
                'sWebEMail = ""
                'While Not rst_Email.EOF
                '    DoEvents
                '    If imTerminate Then
                '        Exit Function
                '    End If
                '    sWebEMail = sWebEMail & gFixQuote(Trim$(rst_Email!emtEmail))
                '    sWebEMail = sWebEMail & ","
                '    rst_Email.MoveNext
                'Wend
        
                SQLQuery = "SELECT arttCode, arttEmail"
                SQLQuery = SQLQuery + " FROM artt"
                SQLQuery = SQLQuery + " WHERE (arttShttCode = " & rst!shttCode
                SQLQuery = SQLQuery + " AND arttWebEMail = 'Y'" & ")"
                Set rst_Email = gSQLSelectCall(SQLQuery)
                    
                sWebEMail = ""
                While Not rst_Email.EOF
                    DoEvents
                    If imTerminate Then
                        Exit Function
                    End If
                    'Remove duplicate references
                    blArttFd = False

                    'D.S. 05/28/13  Commented out For statement below.  The result of it executing is that only
                    'the last date gets shown. i.e. if the dates are 4/1/13, 4/8/13 and 4/15/13 only the 4/15
                    'will be shown.
                    'For llArtt = 0 To UBound(lmArttCode) - 1 Step 1
                    '    If lmArttCode(llArtt) = rst_Email!arttCode Then
                    '        blArttFd = True
                    '        Exit For
                    '    End If
                    'Next llArtt
                    '9960 stop blank email addresses which destroy the previous email address
'                    If Not blArttFd Then
'                        sWebEMail = sWebEMail & gFixQuote(Trim$(rst_Email!arttEmail))
'                        sWebEMail = sWebEMail & ","
'                        lmArttCode(UBound(lmArttCode)) = rst_Email!arttCode
'                        ReDim Preserve lmArttCode(0 To UBound(lmArttCode) + 1) As Long
'                    End If
                    If Not blArttFd Then
                        slTempEmail = gFixQuote(Trim$(rst_Email!arttEmail))
                        If Len(slTempEmail) > 0 Then
                            sWebEMail = sWebEMail & slTempEmail & ","
                            lmArttCode(UBound(lmArttCode)) = rst_Email!arttCode
                            ReDim Preserve lmArttCode(0 To UBound(lmArttCode) + 1) As Long
                        End If
                    End If
                    rst_Email.MoveNext
                Wend
        
                ilLen = Len(sWebEMail)
                If ilLen > 0 Then
                    sWebEMail = Left(sWebEMail, ilLen - 1)
                End If
                
                sWebPW = Trim$(rst!shttWebPW)
                
                'If Len(sWebEMail) = 0 Then
                '    sWebEMail = Trim$(rst!shttWebEmail)
                '    sWebPW = Trim$(rst!ShttWebPW)
                'End If
                '9960
                ilRet = gTestForMultipleEmail(sWebEMail, "RegEmail")
                If ilRet = False Then
                    gMsgBox sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Email Address for Station: " & sCallLetters & " running on: " & smVefName & " Before Continuing", vbExclamation
                    gLogMsg sgErrorMsg & Chr(13) & Chr(10) & "Please Correct the Email Address for Station: " & sCallLetters & " running on: " & smVefName & " Before Continuing", ErrorLog, False
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            
                iFound = False
                'sWebEMail = Trim$(rst!attWebEmail)
                'sWebPW = Trim$(rst!attWebPW)
                lDate = DateValue(gAdjYear(rst!CpttStartDate))
                For iIndex = 0 To UBound(tmEMailRef) - 1 Step 1
                    If (StrComp(Trim$(tmEMailRef(iIndex).sVehName), smVefName, vbTextCompare) = 0) And (StrComp(Trim$(tmEMailRef(iIndex).sCallLetters), sCallLetters, vbTextCompare) = 0) And (tmEMailRef(iIndex).lDate = lDate) Then
                        iFound = True
                        Exit For
                    End If
                Next iIndex
                If Not iFound And ilLen > 0 Then
                    iUpper = UBound(tmEMailRef)
                    tmEMailRef(iUpper).sVehName = smVefName
                    tmEMailRef(iUpper).sCallLetters = sCallLetters
                    tmEMailRef(iUpper).sWebEMail = sWebEMail
                    tmEMailRef(iUpper).sWebPW = sWebPW
                    tmEMailRef(iUpper).iVefCode = imVefCode
                    tmEMailRef(iUpper).iShfCode = iShfCode
                    tmEMailRef(iUpper).lDate = DateValue(gAdjYear(rst!CpttStartDate))
                    sDate = Trim$(Str$(lDate))
                    Do While Len(sDate) < 5
                        sDate = "0" & sDate
                    Loop
                    'tmEMailRef(iUpper).sKey = tmEMailRef(iUpper).sVehName & "|" & tmEMailRef(iUpper).sCallLetters & "|" & sDate
                    tmEMailRef(iUpper).sKey = tmEMailRef(iUpper).sCallLetters & "|" & tmEMailRef(iUpper).sVehName & "|" & sDate
                    ReDim Preserve tmEMailRef(0 To iUpper + 1) As EMAILREF
                End If
            End If
            rst.MoveNext
        Wend
    '9926
    ElseIf (rbcEMailType(MISSED).Value) Then
        slstations = mUnresolvedMissedFromWeb()
        'start on 1 to remove header
        For ilCounter = 1 To UBound(slstations) - 1
            DoEvents
            If imTerminate Then
                Exit Function
            End If
            'because arrays from web are surrounded with double quotes
            slStrippedStationName = Trim$(Replace(slstations(ilCounter), """", ""))
            iShfCode = mIsStationSelected(slStrippedStationName)
            If iShfCode > 0 Then
                sWebEMail = mStationEmailsAsString(iShfCode)
                If Len(sWebEMail) = 0 Then
                    gLogMsg "ERROR: " & slStrippedStationName & " has no email address defined.", "WebNoEmailAddress.Txt", False
                    imNoEmailAddr = True
                Else
                    If Not gTestForMultipleEmail(sWebEMail, "RegEmail") Then
                        gMsgBox sgErrorMsg & Chr(13) & Chr(10) & "Please correct the email address for station " & slStrippedStationName & " before continuing", vbExclamation
                        gLogMsg sgErrorMsg & Chr(13) & Chr(10) & "Please correct the email address for station: " & slStrippedStationName & " before continuing", ErrorLog, False
                        Screen.MousePointer = vbDefault
                        Exit Function
                    End If
                    SQLQuery = "Select ShttWebPw from shtt where shttcode = " & iShfCode
                    Set rst = gSQLSelectCall(SQLQuery)
                    sWebPW = Trim$(rst!shttWebPW)
                    mAddToListOfEmails iShfCode, slStrippedStationName, sWebPW, sWebEMail
                End If


            End If
        Next ilCounter
    End If
    mGatherInfo = True
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebEMail-mGatherInfo"
End Function
'9926
Private Function mUnresolvedMissedFromWeb() As String()
    Dim slstations() As String
    Dim slSql As String
    Dim slDate As String
    Dim llCount As Long
    Dim c As Long
    
    slDate = gObtainPrevMonday(gNow)
    '10229 add '' around the 1
    slSql = "select distinct stationname From MissedSpots ms inner join header on ms.attCode = header.attcode where MGsOnWeb = '1' and  pledgestartdate < '" & Format$(slDate, sgSQLDateForm) & "'"
    'if count = 3 0 is header 1..2..station names 3 is blank
    llCount = gExecWebSQLForVendor(slstations, slSql, True)
    'returns none, or an error (possible if client is not v2)
    If llCount < 1 Then
        ReDim slstations(0)
'        'test only  WHEN database has no unresolved missed spots
'        ReDim slstations(3)
'        slstations(1) = "KAAA-FM"
'        slstations(2) = "KABC-FM"
    End If

    mUnresolvedMissedFromWeb = slstations
End Function
'9926
Private Function mStationEmailsAsString(ilShttCode As Integer) As String
    Dim slSql As String
    Dim rst_Email As ADODB.Recordset
    Dim slRet As String
    Dim ilLen As Integer
    Dim slEMail As String
    
    slRet = ""
    slSql = "SELECT arttCode, arttEmail FROM artt  WHERE arttShttCode = " & ilShttCode & " AND arttWebEMail = 'Y'"
    Set rst_Email = gSQLSelectCall(slSql)
        
    While Not rst_Email.EOF
        DoEvents
        slEMail = gFixQuote(Trim$(rst_Email!arttEmail))
        'currently ignoring.  should I trap instead?
        If Len(slEMail) > 0 Then
            slRet = slRet & slEMail & ","
            lmArttCode(UBound(lmArttCode)) = rst_Email!arttCode
            ReDim Preserve lmArttCode(0 To UBound(lmArttCode) + 1) As Long
        End If
        rst_Email.MoveNext
    Wend
    mStationEmailsAsString = gLoseLastLetterIfComma(slRet)

End Function
'9926 10152-added password
Private Sub mAddToListOfEmails(ilShttCode As Integer, slStationName As String, slPassword As String, slEmails As String)
    Dim blFound As Boolean
    Dim ilIndex As Integer
    Dim ilUpper As Integer
    
    blFound = False
    For ilIndex = 0 To UBound(tmEMailRef) - 1 Step 1
        'I removed test of vehicle
        If StrComp(Trim$(tmEMailRef(ilIndex).sCallLetters), slStationName, vbTextCompare) = 0 Then
            blFound = True
            Exit For
        End If
    Next ilIndex
    If Not blFound Then
        ilUpper = UBound(tmEMailRef)
        tmEMailRef(ilUpper).sVehName = ""
        tmEMailRef(ilUpper).sCallLetters = slStationName
        tmEMailRef(ilUpper).sWebEMail = slEmails
        tmEMailRef(ilUpper).sWebPW = slPassword
        tmEMailRef(ilUpper).iVefCode = 0
        tmEMailRef(ilUpper).lDate = 0
        tmEMailRef(ilUpper).iShfCode = ilShttCode
        ReDim Preserve tmEMailRef(0 To ilUpper + 1) As EMAILREF
    End If
End Sub
'9926
Private Function mIsStationSelected(slStationName As String) As Integer
    Dim ilShttCode As Integer
    Dim ilLoop As Integer
    ilShttCode = 0
    If lbcStation.ListCount > 0 Then
        For ilLoop = 0 To lbcStation.ListCount - 1 Step 1
            If lbcStation.Selected(ilLoop) Then
                If Trim$(lbcStation.List(ilLoop)) = slStationName Then
                    ilShttCode = lbcStation.ItemData(ilLoop)
                    Exit For
                End If
            End If
        Next ilLoop
    Else
        ilShttCode = mShttCodeFromCallLetters(slStationName)
    End If
    mIsStationSelected = ilShttCode
End Function
'9926
Private Function mShttCodeFromCallLetters(slStationName As String) As Integer
    Dim ilShttCode As Integer
    Dim slSql As String
    Dim rst_Shtt As ADODB.Recordset
    
    ilShttCode = 0
    slSql = "Select  shttcode from shtt inner join att on shttcode = attShfCode where shttcallLetters = '" & slStationName & "' "
    If ckcCurrStations.Value = vbChecked Then
        slSql = slSql + " AND attOffAir >= '" & Format$(txtActiveDate.Text, sgSQLDateForm) & "' And attDropDate >= '" & Format$(txtActiveDate.Text, sgSQLDateForm) & "' "
    End If
    Set rst_Shtt = gSQLSelectCall(slSql)
    If Not rst_Shtt.EOF Then
        ilShttCode = rst_Shtt!shttCode
    End If
    mShttCodeFromCallLetters = ilShttCode
End Function
'9926
Private Function mDeletePreviousWebMsgFiles(slToFile As String, slToFileBody As String) As Boolean
    Dim blRet As Boolean
    
    blRet = True
On Error GoTo mDeleteErr:
    If gFileExist(slToFile) = FILEEXISTS Then
        Kill slToFile
        If Not blRet Then
            gMsgBox "Unable to Remove File " & slToFile & " error# " & Str$(Err.Number), vbOKOnly
            Exit Function
        End If
    End If
    If gFileExist(slToFileBody) = FILEEXISTS Then
        On Error GoTo mDeleteErr:
        Kill slToFileBody
        If Not blRet Then
            gMsgBox "Unable to Remove File " & slToFileBody & " error# " & Str$(Err.Number), vbOKOnly
            Exit Function
        End If
    End If
    mDeletePreviousWebMsgFiles = blRet
    Exit Function
mDeleteErr:
    blRet = False
    Resume Next
End Function
Private Function mInit() As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    txtActiveDate.Text = Format$(gNow(), sgShowDateForm)
    SQLQuery = "SELECT * FROM Site Where siteCode = 1"
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        '9926
        If rst!siteAllowMGSpots = "N" Then
            rbcEMailType(MISSED).Visible = False
            'move custom over
            rbcEMailType(CUSTOM).Left = rbcEMailType(MISSED).Left
        End If
        If IsNull(rst!siteOMMinDate) Then
            smOMMinDate = ""
        Else
            smOMMinDate = Trim(rst!siteOMMinDate)
        End If
        
        If Len(Trim(rst!siteOMNoWeeks)) = 0 And Len(Trim(smOMMinDate)) = 0 Then
            gMsgBox "Please go to Site options and enter information into the Overdue Messages fields."
            Exit Function
        End If
    
        If Len(Trim$(smOMMinDate)) > 0 Then
            If gIsDate(Trim$(smOMMinDate)) Then
                smOMMinDate = Format$(smOMMinDate, sgSQLDateForm)
            Else
                smOMMinDate = "1970-01-01"
            End If
        Else
            smOMMinDate = "1970-01-01"
        End If
        
        smWelcome = ""
        If IsNull(rst!siteCCEMail) Then
            smCCEMail = ""
        Else
            smCCEMail = Trim$(rst!siteCCEMail)
            txtCCEMail.Text = smCCEMail
        End If
        If rst!siteWMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!siteWMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                smWelcome = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        smPassword = ""
        If rst!sitePMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!sitePMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                smPassword = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        smOverdue = ""
        If rst!siteOMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!siteOMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                smOverdue = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        '9926
        smMissed = ""
        If rst!siteUMCmtCode > 0 Then
            SQLQuery = "SELECT * FROM CMT Where cmtCode = " & rst!siteUMCmtCode
            Set commrst = gSQLSelectCall(SQLQuery)
            If Not commrst.EOF Then
                smMissed = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
            End If
        End If
        smWelcome = Trim$(smWelcome)
        smPassword = Trim$(smPassword)
        smOverdue = Trim$(smOverdue)
        imOMNoWeeks = rst!siteOMNoWeeks
        smSvWelcome = Trim$(smWelcome)
        smSvPassword = Trim$(smPassword)
        smSvOverdue = Trim$(smOverdue)
        '9926
        smMissed = Trim$(smMissed)
        smSvMissed = smMissed
    Else
        rbcEMailType(WELCOME).Enabled = False
        rbcEMailType(Password).Enabled = False
        rbcEMailType(OVERDUE).Enabled = False
        rbcEMailType(CUSTOM).Enabled = False
        rbcEMailType(WELCOMEBYVEHICLE).Enabled = False
        '9926
        rbcEMailType(MISSED).Enabled = False
    End If
    smCustom = ""
    ilRet = gPopVff()
    mInit = True
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmWebEMail-mInit"
    mInit = False
End Function

Private Sub lbcVehiclesMsg_Click()
    Dim ilItem As Integer
    ilItem = lbcVehiclesMsg.ListIndex
    lbcVehiclesMsg_ItemCheck ilItem
End Sub

Private Sub lbcVehiclesMsg_ItemCheck(Item As Integer)
    Dim ilVef As Integer
    
    If bmInClick Then
        Exit Sub
    End If
    bmInClick = True
    
    mRetainWMVInfo False
    If lbcVehiclesMsg.ListIndex < 0 Then
        bmInClick = False
        Exit Sub
    End If
    imCurrentSelectedVehicle = lbcVehiclesMsg.ListIndex
    lbcVehiclesMsg.Selected(lbcVehiclesMsg.ListIndex) = True
    For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
        If tmWMVInfo(ilVef).iVefCode = lbcVehiclesMsg.ItemData(imCurrentSelectedVehicle) Then
            If lbcVehiclesMsg.Selected(imCurrentSelectedVehicle) Then
                edcMessage.Text = Trim$(tmWMVInfo(ilVef).sComment)
            End If
            bmInClick = False
            Exit Sub
        End If
    Next ilVef
    edcMessage.Text = ""
    tmWMVInfo(UBound(tmWMVInfo)).iCmtCode = 0
    tmWMVInfo(UBound(tmWMVInfo)).iVefCode = lbcVehiclesMsg.ItemData(imCurrentSelectedVehicle)
    tmWMVInfo(UBound(tmWMVInfo)).bChg = False
    tmWMVInfo(UBound(tmWMVInfo)).sComment = ""
    ReDim Preserve tmWMVInfo(0 To UBound(tmWMVInfo) + 1) As WMVINFO
    bmInClick = False
End Sub

Private Sub rbcEMailType_Click(Index As Integer)

    mFillStations
    frcSuppress.Visible = False
    If rbcEMailType(Index).Value Then
    
        If Index <> 3 Then
            txtCustSubject.Visible = False
            ckcCustSubject.Visible = False
            cmcSave.Visible = True
        Else
            cmcSave.Visible = False
        End If
        
        If Index <> 4 Then
            If lbcVehicles.SelCount < 2 Then
                chkCombineEmails.Visible = False
                chkCombineEmails.Value = vbUnchecked
            Else
                chkCombineEmails.Visible = True
            End If
        Else
            chkCombineEmails.Visible = False
        End If
        lbcVehiclesMsg.Visible = False
        mResetMessage Index

        Select Case Index
            Case 0
                edcMessage.Text = smWelcome
                txtActiveDate.Visible = True
                ckcCurrStations.Visible = True
            Case 1
                edcMessage.Text = smPassword
                txtActiveDate.Visible = True
                ckcCurrStations.Visible = True
            Case 2
                edcMessage.Text = smOverdue
                frcSuppress.Visible = True
                txtActiveDate.Visible = False
                ckcCurrStations.Visible = False
            Case 3
                edcMessage.Text = smCustom
                txtActiveDate.Visible = True
                ckcCurrStations.Visible = True
                ckcCustSubject.Visible = True
                ckcCustSubject.Value = vbUnchecked
            Case 4
                lbcVehiclesMsg.Visible = True
                txtActiveDate.Visible = True
                ckcCurrStations.Visible = True
                edcMessage.Text = ""
            '9926
            Case 5
                edcMessage.Text = smMissed
                txtActiveDate.Visible = True
                ckcCurrStations.Visible = True
                chkCombineEmails.Visible = False
        End Select
    End If
End Sub

Private Function mValidDate() As Integer
    
    mValidDate = False
    
    If Not gIsDate(txtActiveDate.Text) Then
        If Trim$(txtActiveDate.Text) = "" Then
            gMsgBox "Please Enter a Valid Date"
        Else
            gMsgBox """" & txtActiveDate.Text & """" & " is not a valid date, please enter another date"
            txtActiveDate.Text = ""
        End If
        'txtActiveDate.SetFocus
        Exit Function
    End If
    
    mValidDate = True

End Function

Private Sub tmcLoadMessages_Timer()
    tmcLoadMessages.Enabled = False
    mLoadVehicleMessages
End Sub

Private Sub txtActiveDate_LostFocus()
    
    mFillStations
'    If ckcCurrStations.Value = vbChecked Then
'        Call mValidDate
'    End If
End Sub

Private Sub mBuildVehicleArray()
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim slStr As String
    ReDim imVefArray(0 To 1) As Integer
    imVefArray(0) = imVefCode
    ilVef = gBinarySearchVef(CLng(imVefCode))
    If ilVef <> -1 Then
        If tgVehicleInfo(ilVef).sVehType = "L" Then
            ilVff = gBinarySearchVff(imVefCode)
            If ilVff <> -1 Then
                slStr = Trim$(tgVffInfo(ilVff).sWebName)
                If slStr <> "" Then
                    smVefName = slStr
                    If sgShowByVehType = "Y" Then
                        smVefName = Mid$(smVefName, 3)
                    End If
                End If
            End If
            For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) Step 1
                If tgVehicleInfo(ilVef).iVefCode = imVefCode Then
                    For ilVff = 0 To UBound(tgVffInfo) - 1 Step 1
                        If tgVehicleInfo(ilVef).iCode = tgVffInfo(ilVff).iVefCode Then
                            If tgVffInfo(ilVff).sMergeWeb <> "S" Then
                                imVefArray(UBound(imVefArray)) = tgVehicleInfo(ilVef).iCode
                                ReDim Preserve imVefArray(0 To (UBound(imVefArray) + 1)) As Integer
                                Exit For
                            End If
                        End If
                    Next ilVff
                End If
            Next ilVef
        End If
    End If
End Sub
Private Sub mFillVehicleMsg()
    Dim iLoop As Integer
    lbcVehiclesMsg.Clear
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        lbcVehiclesMsg.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
        lbcVehiclesMsg.ItemData(lbcVehiclesMsg.NewIndex) = tgVehicleInfo(iLoop).iCode
    Next iLoop
End Sub
Private Sub mRetainWMVInfo(blRetainCurrentVehicle As Boolean)
    Dim ilVef As Integer
    
    If imCurrentSelectedVehicle >= 0 Then
        For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
            If tmWMVInfo(ilVef).iVefCode = lbcVehiclesMsg.ItemData(imCurrentSelectedVehicle) Then
                If lbcVehiclesMsg.Selected(imCurrentSelectedVehicle) Then
                    If Trim$(tmWMVInfo(ilVef).sComment) <> Trim$(edcMessage.Text) Then
                        tmWMVInfo(ilVef).sComment = Trim$(edcMessage.Text)
                        tmWMVInfo(ilVef).bChg = True
                        If sgUstWin(11) = "I" Then
                            cmcSave.Enabled = True
                        End If
                    End If
                Else
                    tmWMVInfo(ilVef).sComment = ""
                End If
                Exit For
            End If
        Next ilVef
    End If
    If Not blRetainCurrentVehicle Then
        imCurrentSelectedVehicle = -1
    End If

End Sub

Private Sub mResetMessage(ilIndex As Integer)

    If ilIndex <> 4 Then
        edcMessage.Left = frcMessage.Left + 345
        edcMessage.Width = frcMessage.Width - 2 * 345 - (cmcSave.Width + 120)
    Else
        lbcVehiclesMsg.Left = frcMessage.Left + 345
        lbcVehiclesMsg.Height = edcMessage.Height + 120
        edcMessage.Left = frcMessage.Left + 345 + lbcVehiclesMsg.Width + 240
        edcMessage.Width = frcMessage.Width - edcMessage.Left - 345 - cmcSave.Width
    End If

End Sub

Private Sub mLoadVehicleMessages()
    Dim ilVef As Integer
    bmInClick = True
    imCurrentSelectedVehicle = -1
    mFillVehicleMsg
    ReDim tmWMVInfo(0 To 0) As WMVINFO
    SQLQuery = "SELECT * FROM CMT Where cmtType = " & "'V'"
    Set commrst = gSQLSelectCall(SQLQuery)
    Do While Not commrst.EOF
        bmInClick = True
        tmWMVInfo(UBound(tmWMVInfo)).iCmtCode = commrst!cmtCode
        tmWMVInfo(UBound(tmWMVInfo)).iVefCode = commrst!cmtVefCode
        tmWMVInfo(UBound(tmWMVInfo)).sComment = commrst!cmtPart1 & commrst!cmtPart2 & commrst!cmtPart3 & commrst!cmtPart4
        ReDim Preserve tmWMVInfo(0 To UBound(tmWMVInfo) + 1) As WMVINFO
        For ilVef = 0 To lbcVehiclesMsg.ListCount - 1 Step 1
            If commrst!cmtVefCode = lbcVehiclesMsg.ItemData(ilVef) Then
                lbcVehiclesMsg.Selected(ilVef) = True
                Exit For
            End If
        Next ilVef
        commrst.MoveNext
    Loop
    bmInClick = False
    lbcVehiclesMsg.ListIndex = -1
End Sub

'10657 rewrote
Private Function mCreateAndFTPFiles() As Boolean
    'welcome by Vehicle loops here by vehicle
    Dim blRet As Boolean
    Dim slErrorMessage As String
    blRet = mCreateFiles()
    If blRet Then
         blRet = mFTPFiles()
    End If
    If Not blRet Then
        'not much in the message department
        slErrorMessage = "E-Mail Vehicle/Station Info"
        If imTerminate Then
            slErrorMessage = "User Terminated: " & slErrorMessage
        Else
            slErrorMessage = "FAILED: " & slErrorMessage
        End If
        SetResults slErrorMessage, MESSAGERED
        gLogMsg slErrorMessage, ErrorLog, False
    End If
    mCreateAndFTPFiles = blRet
End Function
Private Function mSaveMessage(slMessage As String, ilCmtCode As Integer, slType As String, ilVefCode As Integer) As Integer
    Dim slPart1 As String
    Dim slPart2 As String
    Dim slPart3 As String
    Dim slPart4 As String
    On Error GoTo ErrHand
    
    slPart1 = gFixQuote(Mid(slMessage, 1, 255))
    slPart2 = gFixQuote(Mid(slMessage, 256, 255))
    slPart3 = gFixQuote(Mid(slMessage, 511, 255))
    slPart4 = gFixQuote(Mid(slMessage, 766, 255))
    If Len(Trim$(slMessage)) > 0 Then
        If ilCmtCode > 0 Then
            SQLQuery = "Update Cmt Set "
            SQLQuery = SQLQuery & "cmtPart1 = '" & slPart1 & "', "
            SQLQuery = SQLQuery & "cmtPart2 = '" & slPart2 & "', "
            SQLQuery = SQLQuery & "cmtPart3 = '" & slPart3 & "', "
            SQLQuery = SQLQuery & "cmtPart4 = '" & slPart4 & "' "
            SQLQuery = SQLQuery & "Where cmtCode = " & ilCmtCode
            cnn.BeginTrans
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "frmWebEMail-mSaveMessage"
                cnn.RollbackTrans
                mSaveMessage = False
                Exit Function
            End If
            cnn.CommitTrans
        Else
            SQLQuery = "INSERT INTO cmt (cmtType, cmtVefCode, cmtPart1, cmtPart2, cmtPart3, cmtPart4)"
            SQLQuery = SQLQuery & " VALUES ('" & slType & "', " & ilVefCode & ", '" & slPart1 & "', '" & slPart2 & "', "
            SQLQuery = SQLQuery & "'" & slPart3 & "', '" & slPart4 & "')"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "frmWebEMail-mSaveMessage"
                mSaveMessage = False
                Exit Function
            End If
            SQLQuery = "Select MAX(cmtCode) from cmt"
            Set commrst = gSQLSelectCall(SQLQuery)
            ilCmtCode = commrst(0).Value
        End If
    End If
    mSaveMessage = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmfrmWebEMail-mSaveMessage"
    mSaveMessage = False
End Function
Private Function mSetEmailTypeAndFileNames() As Boolean
    Dim blRet As Boolean
    Dim elType As EMAILTYPE
    Dim iLoop As Integer
    Dim ilVef As Integer
    
    blRet = True
    elType = NOTYPE
    If rbcEMailType(WELCOME).Value Then
        elType = WELCOMETYPE
        smToFile = sgExportDirectory & "WebMsg_Welcome.txt"
        smToFileBody = sgExportDirectory & "WebMsg_WelcomeBody.txt"
        smMsgType = "Welcome"
        If Len(Trim(smWelcome)) = 0 Then
            blRet = False
        End If
    ElseIf rbcEMailType(Password).Value Then
        elType = PASSWORDTYPE
        smToFile = sgExportDirectory & "WebMsg_Password.txt"
        smToFileBody = sgExportDirectory & "WebMsg_PasswordBody.txt"
        smMsgType = "Password"
        If Len(Trim(smPassword)) = 0 Then
            blRet = False
        End If
    ElseIf rbcEMailType(OVERDUE).Value Then
        elType = OVERDUETYPE
        smToFile = sgExportDirectory & "WebMsg_Overdue.txt"
        smToFileBody = sgExportDirectory & "WebMsg_OverdueBody.txt"
        smMsgType = "Overdue"
        If Len(Trim(smOverdue)) = 0 Then
            blRet = False
        End If
    ElseIf rbcEMailType(CUSTOM).Value Then
        elType = CUSTOMTYPE
        smToFile = sgExportDirectory & "WebMsg_Custom.txt"
        smToFileBody = sgExportDirectory & "WebMsg_CustomBody.txt"
        smMsgType = "Custom"
        If Len(Trim(smCustom)) = 0 Then
            blRet = False
        End If
    ElseIf rbcEMailType(WELCOMEBYVEHICLE).Value Then
        elType = WELCOMEBYVEHICLETYPE
        smToFile = sgExportDirectory & "WebMsg_Welcome.txt"
        smToFileBody = sgExportDirectory & "WebMsg_WelcomeBody.txt"
        smMsgType = "Welcome"
        chkCombineEmails.Value = vbUnchecked
        mRetainWMVInfo False
        If Len(Trim(smWelcome)) = 0 Then
            For iLoop = 0 To lbcVehicles.ListCount - 1
                If lbcVehicles.Selected(iLoop) Then
                    'Get hmTo handle
                    imVefCode = lbcVehicles.ItemData(iLoop)
                    For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
                        If tmWMVInfo(ilVef).iVefCode = imVefCode Then
                            If Trim$(tmWMVInfo(ilVef).sComment) = "" Then
                                blRet = False
                                Exit For
                            End If
                        End If
                    Next ilVef
                    If Not blRet Then
                        Exit For
                    End If
                End If
            Next iLoop
        End If
    ElseIf rbcEMailType(MISSED).Value Then
        elType = MISSEDTYPE
        smToFile = sgExportDirectory & "WebMsg_Missed.txt"
        smToFileBody = sgExportDirectory & "WebMsg_MissedBody.txt"
        smMsgType = "Missed"
        '10657
        If Len(Trim(smMissed)) = 0 Then
            blRet = False
        End If
    End If
    emEmailType = elType
    If elType = NOTYPE Then
        blRet = False
    End If
    mSetEmailTypeAndFileNames = blRet
End Function
Private Function mLineForFile(ilCounter As Integer, blIncludeVehicle As Boolean, Optional blExcludeDates = False) As String
    Dim slString As String
    
    If blIncludeVehicle Then
        slString = """" & Trim$(gFixDoubleQuote(tmEMailRef(ilCounter).sVehName)) & """" & ","
    End If
    '"vehicle,station,emailaddress,ccemail,"
    slString = slString & """" & Trim$(tmEMailRef(ilCounter).sCallLetters) & """" & "," & """" & Trim$(tmEMailRef(ilCounter).sWebEMail) & """" & "," & """" & smCCEMail & """" & ","
    Select Case emEmailType
        Case MISSEDTYPE
            '"station,emailaddress,ccemail, password"
            slString = """" & Trim$(tmEMailRef(ilCounter).sCallLetters) & """" & "," & """" & Trim$(tmEMailRef(ilCounter).sWebEMail) & """" & "," & """" & smCCEMail & """" & "," & """" & Trim$(tmEMailRef(ilCounter).sWebPW) & """"
        Case OVERDUETYPE
            '"vehicle,station,emailaddress,ccemail, dates"
            If blExcludeDates = False Then
                'if want space before dates
                'slString = slString & " " & """" & Format(tmEMailRef(ilCounter).lDate, "mm/dd/yyyy")
                slString = slString & """" & Format(tmEMailRef(ilCounter).lDate, "mm/dd/yyyy")
            End If
        Case WELCOMETYPE, WELCOMEBYVEHICLE, PASSWORDTYPE
             '"vehicle,station,emailaddress,ccemail, password"
            slString = slString & """" & Trim$(tmEMailRef(ilCounter).sWebPW) & """"
        Case CUSTOMTYPE
            '"vehicle,station,emailaddress,ccemail,custsubject"
            If ckcCustSubject.Value Then
                slString = slString & """" & gFixQuote(Trim$(txtCustSubject.Text)) & """"
            Else
                slString = slString & """" & """"
            End If
    End Select
    mLineForFile = slString
End Function
Private Sub mBasicAndHeaderForFile()
    Dim slStr As String
    Dim ilVef As Integer
    
    Select Case emEmailType
        Case WELCOMETYPE
            Print #hmTo, "vehicle,station,emailaddress,ccemail, password"
            Print #hmToBody, "WelcomeBody"
            Print #hmToBody, """" & smWelcome & """"
        Case PASSWORDTYPE
            Print #hmTo, "vehicle,station,emailaddress,ccemail, password"
            Print #hmToBody, "PaswordBody"
            Print #hmToBody, """" & smPassword & """"
        Case OVERDUETYPE
            Print #hmTo, "vehicle,station,emailaddress,ccemail, dates"
            Print #hmToBody, "OverdueBody"
            Print #hmToBody, """" & smOverdue & """"
        Case CUSTOMTYPE
            Print #hmTo, "vehicle,station,emailaddress,ccemail,custsubject"
            Print #hmToBody, "CustomBody"
            Print #hmToBody, """" & smCustom & """"
        Case MISSEDTYPE
            Print #hmTo, "station,emailaddress,ccemail, password"
            Print #hmToBody, "MissedBody"
            Print #hmToBody, """" & smMissed & """"
        Case WELCOMEBYVEHICLE
            Print #hmTo, "vehicle,station,emailaddress,ccemail, password"
            slStr = smWelcome
            For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
                If tmWMVInfo(ilVef).iVefCode = imVefCode Then
                    slStr = Trim$(tmWMVInfo(ilVef).sComment)
                    Exit For
                End If
            Next ilVef
            Print #hmToBody, """" & slStr & """"
    End Select
End Sub
'TODO can be removed
'Private Function mMessageExistsForType() As Boolean
'    Dim blRet As Boolean
'    Dim iLoop As Integer
'    Dim ilVef As Integer
'    blRet = True
'    Select Case emEmailType
'        Case WELCOMETYPE
'            If Len(Trim(smWelcome)) = 0 Then
'                blRet = False
'            End If
'        Case PASSWORDTYPE
'            If Len(Trim(smPassword)) = 0 Then
'                blRet = False
'            End If
'        Case OVERDUETYPE
'            If Len(Trim(smOverdue)) = 0 Then
'                blRet = False
'            End If
'        Case CUSTOMTYPE
'            If Len(Trim(smCustom)) = 0 Then
'                blRet = False
'            End If
'        Case MISSEDTYPE
'            If Len(Trim(smCustom)) = 0 Then
'                blRet = False
'            End If
'        Case WELCOMEBYVEHICLETYPE
'            chkCombineEmails.Value = vbUnchecked
'            mRetainWMVInfo False
'            If Len(Trim(smWelcome)) = 0 Then
'                For iLoop = 0 To lbcVehicles.ListCount - 1
'                    If lbcVehicles.Selected(iLoop) Then
'                        'Get hmTo handle
'                        imVefCode = lbcVehicles.ItemData(iLoop)
'                        For ilVef = 0 To UBound(tmWMVInfo) - 1 Step 1
'                            If tmWMVInfo(ilVef).iVefCode = imVefCode Then
'                                If Trim$(tmWMVInfo(ilVef).sComment) = "" Then
'                                    blRet = False
'                                    Exit For
'                                End If
'                            End If
'                        Next ilVef
'                        If Not blRet Then
'                            Exit For
'                        End If
'                    End If
'                Next iLoop
'            End If
'    End Select
'    mMessageExistsForType = blRet
'End Function
