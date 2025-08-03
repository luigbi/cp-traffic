VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Begin VB.Form ExportProj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Corporate Export"
   ClientHeight    =   6075
   ClientLeft      =   4245
   ClientTop       =   3630
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "ExportProj.frx":0000
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7560
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
      Height          =   6075
      Left            =   0
      Picture         =   "ExportProj.frx":08CA
      ScaleHeight     =   6015
      ScaleWidth      =   7485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7545
      Begin VB.Timer tmcSetTime 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   1095
         Top             =   5385
      End
      Begin VB.PictureBox plcTo 
         Height          =   375
         Left            =   1185
         ScaleHeight     =   315
         ScaleWidth      =   4245
         TabIndex        =   11
         Top             =   3300
         Width           =   4305
         Begin VB.TextBox edcTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   15
            TabIndex        =   12
            Top             =   30
            Width           =   4230
         End
      End
      Begin VB.CommandButton cmcTo 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
         Height          =   285
         Left            =   5745
         TabIndex        =   10
         Top             =   3420
         Width           =   1485
      End
      Begin VB.Timer tmcClick 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1680
         Top             =   5085
      End
      Begin VB.TextBox edcContract 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   1425
         MaxLength       =   9
         TabIndex        =   9
         Top             =   2160
         Width           =   1200
      End
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   1920
         ItemData        =   "ExportProj.frx":A217
         Left            =   2985
         List            =   "ExportProj.frx":A219
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         Top             =   1140
         Width           =   4215
      End
      Begin VB.CheckBox ckcAll 
         BackColor       =   &H00FFFFC0&
         Caption         =   "All Vehicles"
         Height          =   210
         Left            =   2985
         TabIndex        =   7
         Top             =   840
         Width           =   195
      End
      Begin VB.TextBox edcStart 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   6
         Top             =   810
         Width           =   1200
      End
      Begin VB.TextBox edcEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1485
         Width           =   1200
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
      Begin MSComDlg.CommonDialog CMDialogBox 
         Left            =   135
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   4100
         FontSize        =   0
         MaxFileSize     =   256
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   705
         Index           =   0
         Left            =   585
         TabIndex        =   19
         Top             =   4020
         Visible         =   0   'False
         Width           =   6255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   585
         TabIndex        =   18
         Top             =   3780
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Label lacAll 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "All Vehicles"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3240
         TabIndex        =   17
         Top             =   825
         Width           =   1065
      End
      Begin DZACTXLibCtl.dzactxctrl zpcDZip 
         Left            =   705
         OleObjectBlob   =   "ExportProj.frx":A21B
         Top             =   4710
      End
      Begin VB.Label lacSaveIn 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Save In"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   225
         TabIndex        =   16
         Top             =   3420
         Width           =   810
      End
      Begin VB.Label lacContract 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Contract #"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         TabIndex        =   15
         Top             =   2220
         Width           =   1065
      End
      Begin VB.Label lacStart 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         TabIndex        =   14
         Top             =   870
         Width           =   1065
      End
      Begin VB.Label lacEnd 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         TabIndex        =   13
         Top             =   1545
         Width           =   1065
      End
      Begin VB.Label lacStop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3075
         TabIndex        =   4
         Top             =   5355
         Width           =   1335
      End
      Begin VB.Label lacMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimize"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3075
         TabIndex        =   3
         Top             =   4845
         Width           =   1350
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   60
         Top             =   60
         Width           =   3210
      End
      Begin VB.Label imcOutline 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   420
         Left            =   4785
         TabIndex        =   2
         Top             =   4875
         Visible         =   0   'False
         Width           =   1410
      End
   End
End
Attribute VB_Name = "ExportProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExportProj.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text


'Dim hmMsg As Integer
Dim hmProj As Integer
Dim tmProjByDay() As PROJINFO
Dim tmTradeCnts() As Long       'contract code of trade contracts (partial or full)

Dim smExportName As String
Dim smZipPath As String         'zipped path (base is from revenueexportpath in traffic.ini
Dim imFirstActivate As Integer
Dim lmCntrNo As Long    'for debugging purposes to filter a single contract
Dim lmCntrCode As Long  'contract code so header doesnt have to be read to match all spots

Dim imSetAll As Integer
Dim imAllClicked As Integer

Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image

Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim tmClfSrchKey As CLFKEY0

Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length

Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer       'VSF record length

Dim hmSdf As Integer            'Spot file handle
Dim tmSdf As SDF                'Spot detail record image
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdfSrchKey3 As LONGKEY0

Dim hmSmf As Integer            'Spot MG file handle
Dim tmSmf As SMF                'Spot MG record image
Dim imSmfRecLen As Integer        'Spt MG record length

Dim hmSsf As Integer
Dim tmSsf As SSF                'Spot summary record image
Dim imSsfRecLen As Integer        'SSF record length
Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2 'SSF key record image
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

Dim hmLcf As Integer
Dim tmLcf As LCF                'Log calendar  record image
Dim imLcfRecLen As Integer        'Log calendar record length

Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim imSafRecLen As Integer

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim lmNowDate As Long
Dim lmLastYearStartDate As Long    'earliest date to gather (jan 1 of previous year from todays date)
Dim lmUserStartDate As Long         'user entered start date
Dim lmUserEndDate As Long           'user entered end date
Dim imAutoRun As Integer            '1 = auto run flag, 0 = manual
Dim imClosed As Integer

Private Sub edcTo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
'    gShowBranner
End Sub


Private Sub Form_Load()
    If App.PrevInstance Then
        If imAutoRun = 0 Then
            MsgBox "Only one copy of Corporate Export can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        End If
        gLogMsg "Second copy of Corporate Export path: " & App.Path & " from " & Trim$(gGetComputerName()), "CorporateExport.txt", False
        End
    End If
    gSetBkgdMode
    'igBkgdProg = 6
    mInit
    If imAutoRun = 0 Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMinimized
        gOpenTmf
        tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
        tmcSetTime.Enabled = True
    End If
    tmcClick.Interval = 2000    '2 seconds
    tmcClick.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRes As Integer
    Dim ilRet As Integer
    If imClosed = True Then
        Exit Sub
    End If
    Cancel = 0
    tmcClick.Enabled = False
    
    'If imAutoRun = 0 Then
    '    ilRet = MsgBox("Stop the Corporate Export", vbQuestion + vbYesNo, "Stop Service")
    '    If ilRet = vbNo Then
    '        Cancel = 1
    '        imTerminate = False
    '        Exit Sub
    '    End If
    'End If
    imClosed = True
    imTerminate = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        Me.Left = Screen.Width / 2 - Me.Width / 2
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    If imAutoRun = 1 Then         '9-07-05 if auto run, activate the task
        tmcSetTime.Enabled = False
        gCloseTmf
    End If
    
    Erase tmTradeCnts
    Erase tmProjByDay
    
    btrStopAppl
    Set ExportProj = Nothing   'Remove data segment
    End
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    slCommand = Command$
    sgCommandStr = "debug"
    igDirectCall = 0
    sgIniPath = ""
    'slCommand = sgCommandStr  'Command$
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (StrComp(slCommand, "CS", 1) = 0) Then
        igStdAloneMode = True 'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
    '    ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
    '    If Trim$(slStr) = "" Then
    '        MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
    '        End
    '    End If
    '    ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
    '    ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
    '    If StrComp(slTestSystem, "Test", 1) = 0 Then
    '        ilTestSystem = True
    '    Else
    '        ilTestSystem = False
    '    End If
    '    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    ilPos = InStr(1, slCommand, "/IniLoc:", vbTextCompare)
    If ilPos > 0 Then
        igDirectCall = -1
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8))
        Else
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8, ilSpace - ilPos - 8))
        End If
        If right$(sgIniPath, 1) <> "\" Then
            sgIniPath = sgIniPath & "\"
        End If
    End If
    imAutoRun = 1
    ilPos = InStr(1, slCommand, "/UserInput", 1)
    If ilPos > 0 Then
        imAutoRun = 0
        igBkgdProg = 0
        'gLogMsg "Current Mode: UserInput.", "Corporate Export.txt", False
    Else
        igBkgdProg = 8
       ' gLogMsg "Current Mode: Background.", "Corporate Export.txt", False
    End If
    gInitStdAlone ExportProj, slStr, ilTestSystem
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "ExportProj.Frm"
    If igBkgdProg = 0 Then
        gLogMsg "Current Mode: UserInput.", "Corporate Export.txt", False
    Else
        gLogMsg "Current Mode: Background.", "Corporate Export.txt", False
    End If
End Sub



Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Chf.Btr", Len(tmChf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Clf.Btr", Len(tmClf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cff.Btr", Len(tmCff)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Sdf.Btr", Len(tmSdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Lcf.Btr", Len(tmLcf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Smf.Btr", Len(tmSmf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Vef.Btr", Len(tmVef)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Vsf.Btr", Len(tmVsf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function

Private Sub lacAll_Click()
    If ckcAll.Value = vbChecked Then
        ckcAll.Value = vbUnchecked
    Else
        ckcAll.Value = vbChecked
    End If
End Sub

Private Sub lacMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacMin.Left - 30, lacMin.Top - 45
    imcOutline.Visible = True
End Sub

Private Sub lacStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacStop.Left - 30, lacStop.Top - 45
    imcOutline.Visible = True
End Sub

Private Sub pbcClickFocus_Click()
    imcOutline.Visible = False
End Sub

Private Sub plcSignon_Click()
    imcOutline.Visible = False
End Sub


Private Sub lacMin_Click()
    Dim ilRet As Integer
    If imAutoRun = 0 Then
        If Not imExporting Then
            Screen.MousePointer = vbHourglass
            ilRet = mExport()
            Screen.MousePointer = vbDefault
            If ilRet Then
                lacMin.Caption = "Done"
                lacMin.Enabled = False
            End If
        End If
    Else
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub lacStop_Click()
    If Not imExporting Then
        imTerminate = True
        mTerminate
    Else
        imTerminate = True
    End If
End Sub


Private Sub mBtrErrorMsg(slMsg As String, ilInRet As Integer)
    Dim ilRet As Integer
    
    ilRet = ilInRet
    If ilRet >= 30000 Then
        ilRet = csiHandleValue(0, 7)
    End If
    gLogMsg slMsg & Trim$(Str$(ilRet)), "Corporate Export.txt", False
End Sub

'' **************************************************************************************
''
''  Procedure:  initZIPCmdStruct()
''
''  Purpose:  Set the ZIP control values
''
'' **************************************************************************************
'Sub initZIPCmdStruct()
'  zpcDZip.ActionDZ = NO_ACTION
'  zpcDZip.AddCommentFlag = False
'  zpcDZip.AfterDateFlag = False
'  zpcDZip.BackgroundProcessFlag = False
'  zpcDZip.Comment = ""
'  zpcDZip.CompressionFactor = 5
'  zpcDZip.ConvertLFtoCRLFFlag = False
'  zpcDZip.Date = ""
'  zpcDZip.DeleteOriginalFlag = False
'  zpcDZip.DiagnosticFlag = False
'  zpcDZip.DontCompressTheseSuffixesFlag = False
'  zpcDZip.DosifyFlag = False
'  zpcDZip.EncryptCode = ""  'gCreatePassword
'  zpcDZip.EncryptFlag = False   'True
'  'zpcDZip.ExcludeFollowing = ""
'  'zpcDZip.ExcludeFollowingFlag = False
'  zpcDZip.FixFlag = False
'  zpcDZip.FixHarderFlag = False
'  zpcDZip.GrowExistingFlag = False
'  zpcDZip.IncludeFollowing = ""
'  zpcDZip.IncludeOnlyFollowingFlag = False
'  zpcDZip.IncludeSysandHiddenFlag = False
'  zpcDZip.IncludeVolumeFlag = False
'  zpcDZip.ItemList = ""
'  zpcDZip.MajorStatusFlag = True
'  zpcDZip.MessageCallbackFlag = True
'  zpcDZip.MinorStatusFlag = True
'  zpcDZip.MultiVolumeControl = 0
'
'  'Changed both of these to False from the default True
'  zpcDZip.NoDirectoryEntriesFlag = True
'  zpcDZip.NoDirectoryNamesFlag = True
'
'  zpcDZip.OldAsLatestFlag = False
'  zpcDZip.PathForTempFlag = False
'  zpcDZip.QuietFlag = False
'  zpcDZip.RecurseFlag = False
'  zpcDZip.StoreSuffixes = ""
'  zpcDZip.TempPath = ""
'  zpcDZip.ZIPFile = ""
'
'  'Write out a log file in the windows sub directory
'  zpcDZip.ZipSubOptions = 256
'
'  ' added for rev 3.00
'  zpcDZip.RenameCallbackFlag = False
'  zpcDZip.ExtProgTitle = ""
'  zpcDZip.ZIPString = ""
'
'End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCost                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Cost from line             *
'*                                                     *
'*******************************************************
Private Function mGetCost(tlSdf As SDF, hlClf As Integer, hlCff As Integer, hlSmf As Integer, hlVef As Integer, hlVsf As Integer) As Long
    Dim ilRet As Integer
    Dim slPrice As String
    mGetCost = 0
    If tlSdf.sSpotType = "X" Then
        Exit Function
    End If
    DoEvents
    imClfRecLen = Len(tmClf)
    tmClfSrchKey.lChfCode = tlSdf.lChfCode
    tmClfSrchKey.iLine = tlSdf.iLineNo
    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest Revision
    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
    ilRet = btrGetGreaterOrEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) And ((tmClf.sSchStatus <> "I") And (tmClf.sSchStatus <> "F"))
        DoEvents
        ilRet = btrGetNext(hlClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) And ((tmClf.sSchStatus = "I") Or (tmClf.sSchStatus = "F")) Then
        ilRet = gGetSpotPrice(tlSdf, tmClf, hlCff, hlSmf, hlVef, hlVsf, slPrice)
        If InStr(slPrice, ".") > 0 Then
            mGetCost = gStrDecToLong(slPrice, 2)
        End If
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:8/17/05       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box for conventional           *
'*                      and selling vehicles           *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVpf As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    
    ilRet = gPopUserVehicleBox(ExportProj, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)

    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExportProj
        On Error GoTo 0
    End If
    
    For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
        slNameCode = tgUserVehicle(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        ilVpf = gBinarySearchVpf(ilVefCode)
        If ilVpf <> -1 Then
            If tgVpf(ilVpf).sExpHiCorp = "Y" Then
                lbcVehicle.Selected(ilLoop) = True
            End If
        End If
    Next ilLoop
    
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'********************************************************************************************
'           mCrProj- The Export gathers all spots for approx 24 months.
'                    Goes back to previous year starting jan 1 and
'                    continues to the future for as long as there are spots sched.
'                    A text file is created with inventory, avails, spot counts,
'                    revenue counts by hour within day and spot length by vehicle
'
'           Process by vehicle for all days, hours, spot lengths
'           11-15-05  selective contracts was never tested
'                     CCT wants to include all trades (partials & full) and
'                     always exclude fill spots
'********************************************************************************************
Function mCrProj() As Integer
    Dim ilRet As Integer
    Dim ilError As Integer
    
    Dim ilIndex As Integer
    ReDim ilSpotLens(0 To 0) As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim llEarliestLcfDate As Long
    Dim llLatestLcfDate As Long
    Dim llDate As Long
    Dim ilDate(0 To 1) As Integer
    Dim llTime As Long
    Dim ilWhichHour As Integer
    Dim ilSpot As Integer
    Dim slDate As String
    Dim ilType As Integer
    ReDim tlSdf(0 To 0) As SDF
    ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT

    ilError = 0             'assume everything is OK
    ilIndex = gBinarySearchVpf(tmVef.iCode)
    If ilIndex <> -1 Then
        mSortLengths ilIndex, ilSpotLens()      'sort the spot lengths, to be used
           
        
        If lmUserStartDate = 0 Then     'if no date entered, determine the earliest from the log calendar
            llEarliestLcfDate = gGetEarliestLCFDate(hmLcf, "C", tmVef.iCode)
            If llEarliestLcfDate > lmLastYearStartDate Then     'check that the start date of last year has log calendar in effect,
                                                                'if not, use the earliest log calendar date
                lmLastYearStartDate = llEarliestLcfDate
            End If
        Else                            'user entered a date
            lmLastYearStartDate = lmUserStartDate
        End If
        
        If lmUserEndDate = 0 Then       'if no date entered, determine the latest date from the log calendar
            llLatestLcfDate = gGetLatestLCFDate(hmLcf, "C", tmVef.iCode)     'determine how far in the future to obtain stats
        Else
            llLatestLcfDate = lmUserEndDate
        End If
        
        'show Processing vehicle & dates on caption screen & file
        gLogMsg "Processing " & Trim$(tmVef.sName) & " for " & Format(lmLastYearStartDate, "m/d/yy") & " - " & Format(llLatestLcfDate, "m/d/yy"), "Corporate Export.txt", False
        lacInfo(1).Caption = "Processing " & Trim$(tmVef.sName) & " for " & Format(lmLastYearStartDate, "m/d/yy") & " - " & Format(llLatestLcfDate, "m/d/yy")
        lacInfo(1).Visible = True
        ilType = 0
        'use the ssf for inventory only
        For llDate = lmLastYearStartDate To llLatestLcfDate
            DoEvents
            'Create arrays for each spot length for this day- all stats will be built in them
            ilUpper = 0
            ReDim tmProjByDay(0 To 0) As PROJINFO              'init for next vehicle
            For ilLoop = 0 To UBound(ilSpotLens) - 1
                tmProjByDay(ilLoop).iLen = ilSpotLens(ilLoop)
                ilUpper = ilUpper + 1
                ReDim Preserve tmProjByDay(0 To ilUpper) As PROJINFO
            Next ilLoop
            
            gPackDateLong llDate, ilDate(0), ilDate(1)
            imSsfRecLen = Len(tmSsf)
            If tmVef.sType <> "G" Then
                tmSsfSrchKey.iType = 0 'slType-On Air
                tmSsfSrchKey.iVefCode = tmVef.iCode
                tmSsfSrchKey.iDate(0) = ilDate(0)
                tmSsfSrchKey.iDate(1) = ilDate(1)
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Else
                tmSsfSrchKey2.iVefCode = tmVef.iCode
                tmSsfSrchKey2.iDate(0) = ilDate(0)
                tmSsfSrchKey2.iDate(1) = ilDate(1)
                ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
                ilType = tmSsf.iType
            End If
            DoEvents
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = tmVef.iCode) And (tmSsf.iDate(0) = ilDate(0)) And (tmSsf.iDate(1) = ilDate(1))
                DoEvents
                For ilLoop = 1 To tmSsf.iCount Step 1
                    DoEvents
                   LSet tmAvail = tmSsf.tPas(ilLoop)
                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        'process the inventory for the designated hour
                        ilWhichHour = (llTime \ 3600) + 1
                        mBuildInvByLength ilWhichHour, ilSpotLens()
                    End If
                Next ilLoop
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tmVef.sType = "G" Then
                    ilType = tmSsf.iType
                End If
            Loop
            slDate = Format$(llDate, "m/d/yy")
            mBuildSpotStats slDate, tlSdf()    'obtain the spots for a single date
            ilRet = mWriteExportRec(slDate)               'create up to 24 records for the day (ignore hours without inventory)
            If Not ilRet Then
                ilError = 2
                Exit For
            End If
        Next llDate
    Else                    'didnt find vehicle options
        ilError = 1
    End If
    
    mCrProj = ilError
End Function

'
'
'
'           mCloseProjectionfiles - Close all applicable files for
'                      Corporate projection Export
'
Sub mCloseProjFiles()
    Dim ilRet As Integer
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmCHF)

    btrDestroy hmVef
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmSdf
    btrDestroy hmSsf
    btrDestroy hmLcf
    btrDestroy hmVsf
    btrDestroy hmSmf
    btrDestroy hmCHF
End Sub

'
'
'           mOpenProjectionFiles - open files applicable to Corporate Projection Export
'                           8-17-05
'
'
Function mOpenProjFiles() As Integer
    Dim ilRet As Integer
    Dim ilTemp As Integer
    Dim ilError As Integer
    Dim slStamp As String

    ilError = False

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen VEF)", ExportProj
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen CHF)", ExportProj
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen CLF)", ExportProj
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen CFF)", ExportProj
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen SDF)", ExportProj
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen SSF)", ExportProj
    On Error GoTo 0
    imSsfRecLen = Len(tmSsf)
    
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen LCF)", ExportProj
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen SMF)", ExportProj
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenProjFilesErr
    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen VSF)", ExportProj
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)


    mOpenProjFiles = ilError
    Exit Function

mOpenProjFilesErr:
    ilError = True
    Return
End Function

'
'
'           mWriteExportRec - gather all the information for a day, hour, and spot length and write
'           a record to the export .txt file
'
'           <input> tlProjInfo - structure containing all the info required to previous years data
'                               starting jan1, and going thru the last spot scheduled for the vehicle.
'                               Create records by hour within date by vehicle and spot length.
'           Return - true if error, otherwise false
Private Function mWriteExportRec(slDate As String) As Integer
    Dim ilLoop As Integer
    Dim slRecord As String
    Dim ilIndex As Integer
    Dim slVehicle As String
    Dim ilError As Integer
    Dim ilRemainder As Integer
    Dim ilRecdLen As Integer
    Dim slMonth As String
    Dim slYear As String
    Dim slAmt As String
    Dim slStr As String
    Dim llAvgRate As Long
    Dim ilWhichHour As Integer
    Dim ilRet As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim slDayDescr As String * 3
    Dim slDayOfWeek As String * 21

    DoEvents
    slDayOfWeek = "MONTUEWEDTHUFRISATSUN"
    ilError = False
   
    'format the month info for a contract/vehicle
    slVehicle = ""
    ilIndex = gBinarySearchVef(tmVef.iCode)
    If ilIndex <> -1 Then
        slVehicle = Trim$(tmVef.sName)
        
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        'obtain the day of the week
        ilDay = gWeekDayStr(slDate)
        slDayDescr = Mid$(slDayOfWeek, (ilDay * 3) + 1, 3)
        For ilLoop = LBound(tmProjByDay) To UBound(tmProjByDay) - 1
            For ilWhichHour = 1 To 24
                DoEvents
                'ignore record if no inventory defined
                If tmProjByDay(ilLoop).iSchedUnits(ilWhichHour) <> 0 Or tmProjByDay(ilLoop).iInventory(ilWhichHour) <> 0 Then        'need at least 1 spot sch or inv to create record
                    slRecord = Trim$(tmVef.sName) & Chr$(9)              'vehicle name & tab
                    'slRecord = slRecord & Val(tmVef.sCodeStn) & Chr$(9)       'this is a string field, user needs an integer so it must
                                                                     'this field should be a number input
                                                                     'dont need this as long as they have call letters and band
                    
                    slRecord = slRecord & slYear & "-" & slMonth & "-" & Trim$(slDay) & " " & Trim$(Str(ilWhichHour - 1)) & ":00:00" & Chr$(9)   'date ,  hour (military) , tab
                    
                    slRecord = slRecord & slDayDescr & Chr$(9)          'day of week descr, tab
                    slRecord = slRecord & Trim$(Str(tmProjByDay(ilLoop).iLen)) & Chr$(9)        'spot length, tab
                    
                    slRecord = slRecord & "0" & Chr$(9)                 'BTA (always 0) , tab
                    slRecord = slRecord & Trim$(Str(tmProjByDay(ilLoop).iInventory(ilWhichHour))) & Chr$(9)  'inventory , tab
                      
                    slAmt = gLongToStrDec(tmProjByDay(ilLoop).lMinRate(ilWhichHour), 2)
                    slRecord = slRecord & slAmt & Chr$(9)           'min $ value , tab
                    
                    slAmt = gLongToStrDec(tmProjByDay(ilLoop).lMaxRate(ilWhichHour), 2)
                    slRecord = slRecord & slAmt & Chr$(9)           'max $ value , tab
                    
                    If tmProjByDay(ilLoop).iSchedUnits(ilWhichHour) <> 0 Then
                        llAvgRate = tmProjByDay(ilLoop).lSchedRev(ilWhichHour) / tmProjByDay(ilLoop).iSchedUnits(ilWhichHour) 'determine avg rate (sched revenue/sched units)
                        slAmt = gLongToStrDec(llAvgRate, 2)
                    Else
                        slAmt = ".00"
                    End If
                    slRecord = slRecord & slAmt & Chr$(9)               'avg unit rate, tab
                    
                    slAmt = gLongToStrDec(tmProjByDay(ilLoop).lSchedRev(ilWhichHour), 2)
                    slRecord = slRecord & slAmt & Chr$(9)           'schedule revenue , tab
                    
                    slAmt = gLongToStrDec(tmProjByDay(ilLoop).lMissedRev(ilWhichHour), 2)
                    slRecord = slRecord & slAmt & Chr$(9)           'missed (pooled) revenue , tab
                    
                    slAmt = Trim$(Str(tmProjByDay(ilLoop).iSchedUnits(ilWhichHour)))
                    slRecord = slRecord & slAmt & Chr$(9)           'scheduled units , tab
                    
                    
                    slAmt = Trim$(Str(tmProjByDay(ilLoop).iNCUnits(ilWhichHour)))
                    slRecord = slRecord & slAmt & Chr$(9)           'zero rate units , tab
                    
                    slAmt = Trim$(Str(tmProjByDay(ilLoop).iMissedUnits(ilWhichHour)))
                    slRecord = slRecord & slAmt & Chr$(9)           'missed (pooled) units , tab
    
                    llAvgRate = tmProjByDay(ilLoop).iInventory(ilWhichHour) - tmProjByDay(ilLoop).iSchedUnits(ilWhichHour)
                    slAmt = Trim$(Str(llAvgRate))
                    slRecord = slRecord & slAmt & Chr$(9)               'avails, tab
                    
                    slRecord = slRecord & "0" & Chr$(9)                 'unused time (always 0), tab
                    slRecord = slRecord & Chr$(13) & Chr$(10)       'carriage return, line Chr$(9)
    
                    ilRecdLen = Len(Trim(slRecord))
                    ilRet = 0
                    On Error GoTo mWriteExportRecErr
                    Print #hmProj, Left(slRecord, ilRecdLen)
                    On Error GoTo 0
                    If ilRet <> 0 Then
                        imExporting = False
                        gLogMsg "Error writing to Export File, Error # " & Str$(ilRet), "Corporate Export.txt", False
                        mWriteExportRec = False
                        Exit Function
                    End If
                End If                          'zero inventory
            Next ilWhichHour
        Next ilLoop

    End If

    mWriteExportRec = True
    Exit Function

mWriteExportRecErr:
    ilRet = err.Number
    Resume Next

End Function


Private Sub ckcAll_Click()
Dim Value As Integer
Dim llValue As Integer
Dim llRg As Long
Dim ilValue As Integer
Dim llRet As Long

    If lbcVehicle.ListCount <= 0 Then
        Exit Sub
    End If

    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If

    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcVehicle.HWnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub


Private Sub cmcTo_Click()
    CMDialogBox.DialogTitle = "Export To File"
    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
    CMDialogBox.DefaultExt = ".Csv"
    CMDialogBox.flags = cdlOFNCreatePrompt
    CMDialogBox.Action = 1 'Open dialog
    edcTo.Text = CMDialogBox.fileName
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
    If edcTo.Text = "" Then
        edcTo.Text = smExportName
    End If
End Sub
Private Sub edcContract_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEnd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStart_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcTo_Change()
    mSetCommands
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slTodayDate As String
    Dim slLastYearDate As String
    Dim slDay As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim slEndDate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVpf As Integer
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer

    mParseCmmdLine
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    If imAutoRun = 0 Then
        lacMin.Caption = "Export"
    Else
        lacMin.Caption = "Minimize"
    End If
    imClosed = False
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imExporting = False
    imBypassFocus = False
    imSetAll = True
    imAllClicked = False
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    
    ilRet = gObtainVef() 'Build into tgMVef
    If ilRet = False Then
        imTerminate = True
    End If
    mVehPop
    
    For ilVehicle = 0 To lbcVehicle.ListCount - 1
        slNameCode = tgUserVehicle(ilVehicle).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        ilVpf = gBinarySearchVpf(ilVefCode)
        If ilVpf <> -1 Then
            If (tgVpf(ilVpf).sExpHiCorp <> "N") Then     'this vehicle flagged to export
                lbcVehicle.Selected(ilVehicle) = True
            End If
        End If
     Next ilVehicle
    
    slTodayDate = Format$(lmNowDate, "m/d/yy")
    gObtainYearMonthDayStr slTodayDate, True, slYear, slMonth, slDay
    slLastYearDate = "1/1/" & Trim$(Str$(Val(slYear) - 1))
    edcStart.Text = slLastYearDate
    lmLastYearStartDate = gDateValue(slLastYearDate)             'determine start of last years date
  
    Screen.MousePointer = vbDefault

    Exit Sub
         
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'
    Dim ilRet As Integer
  
    Screen.MousePointer = vbDefault
    'igParentRestarted = False
    'If Not igStdAloneMode Then
    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
    '    Else
    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '        edcLinkDestHelpMsg.LinkExecute "Done"
    '    End If
    '    Do While Not igParentRestarted
    '        DoEvents
    '    Loop
    'End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Traffic
    Unload ExportProj
    'Set ExportProj = Nothing   'Remve data segment
    igManUnload = NO
End Sub

Private Sub lbcVehicle_Click()
    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked  '9-12-02 False
        imSetAll = True
    End If
    mSetCommands
End Sub


Private Sub tmcClick_Timer()
    Dim slRepeat As String
    Dim ilRet As Integer
    'Dim slDateTime As String
    Dim slDefaultFileName As String

    tmcClick.Enabled = False
    If imTerminate Then
        lacStop_Click
        Exit Sub
    End If
    
    mCheckGG
    If igGGFlag = 0 Then
        imTerminate = True
        lacStop_Click
        Exit Sub
    End If
    
    'Determine name of export (.txt file)
    'for now, do not create unique names that append a new letter of the alphabet.
    'Always ask to overwrite
    slRepeat = ""
    slDefaultFileName = "dbo-tDailyProjection" & Trim$(slRepeat)
    'Do
        ilRet = 0
        On Error GoTo cmcExportDupNameErr:
        smExportName = sgExportPath & slDefaultFileName & ".txt"
        If sgRevenueExportPath = "" Then
            smZipPath = sgExportPath & slDefaultFileName & ".txt"           '11-3-05 append extension
        Else
            smZipPath = sgRevenueExportPath & slDefaultFileName & ".txt"    '11-3-05 append extension
        End If
        'slDateTime = FileDateTime(smExportName)
        ''If ilRet = 0 Then       'fell thru , there was a filename that existed with same name. Increment the letter
        ''    If slRepeat = "" Then       'first time doesnt have the alpha appended for conseutive runs
        ''        slRepeat = "A"
        ''    Else
        ''        slRepeat = Chr(Asc(slRepeat) + 1)
        ''    End If
        ''End If
    'Loop While ilRet = 0
    edcTo.Text = smExportName
    edcTo.Visible = True
    
    If imAutoRun = 1 Then         '9-07-05 if auto run, activate the task
        '9-30-05 changed to use field in vpf (vpf.sExpHiCorp to indicate if  vehicle should be included)
        'ckcAll.Value = vbChecked    'assume all vehicles
        gUpdateTaskMonitor 1, "CE"
        ilRet = mExport()
        gUpdateTaskMonitor 2, "CE"
    End If
    Exit Sub
cmcExportDupNameErr:
    ilRet = 1
    Resume Next
End Sub

Public Sub mSetCommands()

    If lbcVehicle.SelCount > 0 And edcTo.Text <> "" Then
        'cmcExport.Enabled = True
        lacMin.Enabled = True
    Else
        'cmcExport.Enabled = False
        lacMin.Enabled = False
    End If
    
End Sub

'
'               msortLengths - sort the spots lengths stored with
'               vehicle options in descending order.  Used to determine
'               inventory in breaking them out by lengths.  Filter out
'               the items unused (zero lengths)
Private Sub mSortLengths(ilIndex As Integer, ilSpotLens() As Integer)
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilEventLoop As Integer
    Dim ilCounter As Integer
    Dim ilTemp As Integer

    ilUpper = 0
    For ilLoop = 0 To 9
        If tgVpf(ilIndex).iSLen(ilLoop) > 0 Then
            ilSpotLens(ilUpper) = tgVpf(ilIndex).iSLen(ilLoop)
            ilUpper = ilUpper + 1
            ReDim Preserve ilSpotLens(0 To ilUpper)
        End If
    Next ilLoop
    
    For ilEventLoop = 0 To ilUpper - 1
        For ilCounter = 0 To ilUpper - 2

            If ilSpotLens(ilCounter) < ilSpotLens(ilCounter + 1) Then
                ilTemp = ilSpotLens(ilCounter)
                ilSpotLens(ilCounter) = ilSpotLens(ilCounter + 1)
                ilSpotLens(ilCounter + 1) = ilTemp
            End If
        Next ilCounter
    Next ilEventLoop
    Exit Sub
End Sub
'
'               mBuildInvBylength - accumulate the amount of inventory
'               for the avail.  Use the spot length table to determine
'               which lengths to use for inventory.  Inv will be
'               calculated using the highest spot length down to
'               the lowest.  For example:  Vehicle spots lengths: 60 30 15 10
'               A 2M45S avail would result in 2 60 units,1 30 unit, 1 15 unit
'               <input> ilWhichHour - hour of day to build stats (1-24)
'                       ilSpotLens() - array of descending spot lengths from veh options
'               8-18-05
Private Sub mBuildInvByLength(ilWhichHour As Integer, ilSpotLens() As Integer)
    Dim ilAvailLen As Integer
    Dim ilRemaining As Integer
    Dim ilLoop As Integer
    Dim ilUnits As Integer
    Dim ilStats As Integer

    ilRemaining = tmAvail.iLen
    For ilLoop = 0 To UBound(ilSpotLens) - 1
        ilUnits = ilRemaining \ ilSpotLens(ilLoop)
        ilRemaining = ilRemaining - (ilUnits * ilSpotLens(ilLoop))
        'udpate the inventory statistics entry for the matching spot length
        For ilStats = LBound(tmProjByDay) To UBound(tmProjByDay) - 1
            If tmProjByDay(ilStats).iLen = ilSpotLens(ilLoop) Then
                tmProjByDay(ilStats).iInventory(ilWhichHour) = tmProjByDay(ilStats).iInventory(ilWhichHour) + ilUnits
            End If
        Next ilStats
    Next ilLoop
    Exit Sub
End Sub
'
'           mBuildspotStats - determine stats for revenue, spots sched/missed & availability
'           Read the spots by vehicle, date, time & sch status for one day at a time
'           Arrays are maintained by spot length.  Within that spot length array are
'           24 hour buckets representing revenue, inventory, sold, N/c spots, etc.
'           Ignore PSA, Promo & 100% trade contracts
'           11-14-05 CCT wants to include all trades (partials & full trade), and exlude fills
Private Sub mBuildSpotStats(slDate As String, tlSdf() As SDF)
    Dim ilStats As Integer
    Dim llAmt As Long
    Dim ilSpots As Integer
    Dim ilRet As Integer
    Dim ilWhichHour As Integer
    Dim llTime As Long
    Dim ilTrade As Integer
    Dim ilSpotOK As Integer
  
    ilRet = gGetSpotsbyVefDate(hmSdf, tmVef.iCode, slDate, slDate, tlSdf())
    For ilSpots = LBound(tlSdf) To UBound(tlSdf) - 1
        DoEvents
        tmSdf = tlSdf(ilSpots)
        'ignore psa, promo, fills (11-14-05) and hidden spots
        If (tmSdf.sSpotType <> "S" And tmSdf.sSpotType <> "M" And tmSdf.sSpotType <> "X" And tmSdf.sSchStatus <> "H") And (lmCntrCode = 0 Or lmCntrCode = tmSdf.lChfCode) Then        'exclude PSA/Promos, fills and hidden spots
            gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
            'process the scheduled spts for the designated hour
            ilWhichHour = (llTime \ 3600) + 1
            For ilStats = LBound(tmProjByDay) To UBound(tmProjByDay) - 1        'find matching spot length entry
                DoEvents
                If tmProjByDay(ilStats).iLen = tmSdf.iLen Then
                    ilSpotOK = True
                    '11-14-05 trades are always included; this array is never created
                    For ilTrade = 0 To UBound(tmTradeCnts) - 1
                        DoEvents
                        If tmSdf.lChfCode = tmTradeCnts(ilTrade) Then   'if matching contract header code, this is a trade and
                                                                        'should be ignored
                            ilSpotOK = False
                            Exit For
                        End If
                    Next ilTrade
                    If ilSpotOK Then
                        DoEvents
                        If tmSdf.sSpotType = "X" Or tmSdf.sSpotType = "O" Or tmSdf.sSpotType = "C" Then 'no rates apply for fill spots, or bb open/close spots
                            llAmt = 0
                        Else
                            llAmt = mGetCost(tmSdf, hmClf, hmCff, hmSmf, hmVef, hmVsf)
                        End If
                        
                        If llAmt = 0 Then
                            tmProjByDay(ilStats).iNCUnits(ilWhichHour) = tmProjByDay(ilStats).iNCUnits(ilWhichHour) + 1     'total zero rate units
                        End If
    
                        If tmProjByDay(ilStats).iNCUnits(ilWhichHour) = 0 Then           'no NC exists yet, see if this spot is the lowest rate
                            If tmProjByDay(ilStats).lMinRate(ilWhichHour) = 0 Then      'first time, set the lowest rate
                                tmProjByDay(ilStats).lMinRate(ilWhichHour) = llAmt
                            Else                                                        'lowest rate has been set so far, see if this spot is lower
                                If llAmt < tmProjByDay(ilStats).lMinRate(ilWhichHour) Then        'get the lowest price in this hour
                                    tmProjByDay(ilStats).lMinRate(ilWhichHour) = llAmt
                                End If
                            End If
                        End If
                        If llAmt > tmProjByDay(ilStats).lMaxRate(ilWhichHour) Then      'get the Highest price in this hour
                            tmProjByDay(ilStats).lMaxRate(ilWhichHour) = llAmt
                        End If
                        
                        tmProjByDay(ilStats).lSchedRev(ilWhichHour) = tmProjByDay(ilStats).lSchedRev(ilWhichHour) + llAmt   'schedule revenue
                        tmProjByDay(ilStats).iSchedUnits(ilWhichHour) = tmProjByDay(ilStats).iSchedUnits(ilWhichHour) + 1   'scheduled units
                       
                        If tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "M" Then    'missed or cancelled
                            tmProjByDay(ilStats).lMissedRev(ilWhichHour) = tmProjByDay(ilStats).lMissedRev(ilWhichHour) + llAmt   'schedule revenue
                            tmProjByDay(ilStats).iMissedUnits(ilWhichHour) = tmProjByDay(ilStats).iMissedUnits(ilWhichHour) + 1   'schedule revenue
                        End If
                    End If                  'ilspotOK
                End If
            Next ilStats
        End If
    Next ilSpots
    Exit Sub
End Sub


'                   Create Corporate Projection file (tab deliminted) that contains inventory from SSF
'                   and spot statistics from SDF.  Each unique record is broken out by date,
'                   spot length and hour.  Inventory is determined by the max lengths allowed.
'                   For example:  2/60 is treated as 1 60" unit, rather than 2 30s.  If a 30 is
'                   booked, it is considered oversold by 1 30", and open to 1 60" spot.
'                   Revenue is gathered as well as missed spots.
'                   PSAS and Promos are ignored, as well as trade contracts.
'
Private Function mExport()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDateTime As String
    Dim ilYear As Integer
    Dim slStart As String
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilAnswer As Integer
    Dim ilError As Integer
    Dim ilErrorStatus As Integer
    Dim ilVefIndex As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slCntrType As String
    Dim slCntrStatus As String
    Dim ilHOState As Integer
    Dim ilChf As Integer
    Dim ilUpper As Integer
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    ReDim tmTradeCnts(0 To 0) As Long
    Dim szFile As String

    
    lacInfo(0).Visible = False

    If imExporting Then
        Exit Function
    End If

    lmCntrNo = 0                'ths is for debugging on a single contract
    slStr = ExportProj!edcContract
    If slStr <> "" Then
        lmCntrNo = Val(slStr)
    End If

    'smExportFile contains the name to use which has been moved to edcTo.Text
    smExportName = Trim$(edcTo.Text)
    If Len(smExportName) = 0 Then
        Beep
        edcTo.SetFocus
        Exit Function
    End If

    If (InStr(smExportName, ":") = 0) And (Left$(smExportName, 2) <> "\\") Then     'test for absence of colon and not using \\
        smExportName = sgExportPath & smExportName
    End If

    ilRet = 0
    'On Error GoTo mExportErrErr:
    'slDateTime = FileDateTime(smExportName)
    ilRet = gFileExist(smExportName)
    If ilRet = 0 Then
        'always kill the duplicate file name
        
        'file already exists, continue?
        'ilAnswer = MsgBox("Filename already exists, overwrite?", vbYesNo + vbApplicationModal, "Save In")
        'If ilAnswer = vbYes Then
            Kill smExportName
        'Else
        '    Exit Sub
        'End If
    End If
    gLogMsg "", "Corporate Export.txt", True
    gLogMsg "** Corporate Export " & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "Corporate Export.txt", True
    gLogMsg "** Exporting to " & ": " & smExportName, "Corporate Export.txt", False
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo mExportErrErr:
    gLogMsg "** Removing file " & ": " & smExportName, "Corporate Export.txt", False
    'hmProj = FreeFile
    gLogMsg "** Creating file " & ": " & smExportName, "Corporate Export.txt", False
    'Open smExportName For Output As hmProj
    ilRet = gFileOpen(smExportName, "Output", hmProj)
    If ilRet <> 0 Then
        gLogMsg "Open Error #" & Str$(err.Numner) & smExportName, "Corporate Export.txt", False
        Close #hmProj
        imExporting = False
        Screen.MousePointer = vbDefault
        If imAutoRun = 0 Then
            MsgBox "Open Error #" & Str$(err.Numner) & smExportName, vbOKOnly, "Open Error"
        End If
        Exit Function
    End If
    gLogMsg "** Storing Output into " & smExportName & " **", "Corporate Export.txt", False

    Screen.MousePointer = vbHourglass
    imExporting = True

    If mOpenProjFiles() = 0 Then
        ilErrorStatus = 0
        slStartDate = Format$(lmLastYearStartDate, "m/d/yy")
        slStr = Format$(lmNowDate, "m/d/yy")
        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
        slEndDate = "12/31/" & Trim$(slYear)
        
        'test for valid single contract if entered
        If lmCntrNo <> 0 Then
            tmChfSrchKey1.lCntrNo = lmCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_END_OF_FILE Or tmChf.lCntrNo <> lmCntrNo Then
                'exit, no contract found
                lacInfo(0).Caption = "Contract #" & Trim$(ExportProj!edcContract) & " does not exist, application aborted"
                gLogMsg "Contract #" & Trim$(ExportProj!edcContract) & " does not exist, application aborted", "Corporate Export.txt", False
                Close #hmProj
                mCloseProjFiles
                
                lacInfo(0).Visible = True
                lacStop.Caption = "&Done"
                Screen.MousePointer = vbDefault
                imExporting = False
                
                If imAutoRun = 1 Then     '9-7-05 if auto run, task finished, unload and return to caller
                    'cmcCancel_Click
                    gUpdateTaskMonitor 2, "CE"
                    lacStop_Click
                End If
                Exit Function
            Else
                lmCntrCode = tmChf.lCode
            End If
        End If
        
        'build array of trade contrcts so the contract header doesnt have to be reread for each spot
        slCntrStatus = "HOGN"                 'statuses: hold, order, unsch hold, uns order
        slCntrType = "CVTRQ"         'all types: PI, DR, etc.  except PSA(p) and Promo(m)
        ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)
        
        ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
        ilRet = gObtainCntrForDate(ExportProj, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
        ilUpper = 0
        For ilChf = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1
           
            '11-14-05 CCT wants to always include all trades
            'If tlChfAdvtExt(ilChf).iPctTrade = 100 Then         'only 100% trades excluded
            '    tmTradeCnts(ilUpper) = tlChfAdvtExt(ilChf).lCode        'save the contr codes to match the spot records (sdfchfcode)
            '    ilUpper = ilUpper + 1
            '    ReDim Preserve tmTradeCnts(0 To ilUpper) As Long
            'End If
        Next ilChf

        'get user entered dates for testing purposes
        If edcStart.Text = "" Then
            lmUserStartDate = 0
        Else
            lmUserStartDate = gDateValue(edcStart.Text)
        End If
        If edcEnd.Text = "" Then
            lmUserEndDate = 0
        Else
            lmUserEndDate = gDateValue(edcEnd.Text)
        End If
        
        For ilVehicle = 0 To lbcVehicle.ListCount - 1
            DoEvents
            If lbcVehicle.Selected(ilVehicle) = True Then
                slNameCode = tgUserVehicle(ilVehicle).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVefIndex = gBinarySearchVef(ilVefCode)
                tmVef = tgMVef(ilVefIndex)
                DoEvents
                ilError = mCrProj()
                If ilError = 1 Then   'returned with error?
                    lacInfo(0).Caption = Trim$(tmVef.sName) & " Options table missing"
                    gLogMsg " Export error - " & Trim$(tmVef.sName) & " options table missing", "Corporate Export.txt", False
                    ilErrorStatus = 1
                ElseIf ilError = 2 Then
                    ilErrorStatus = ilError
                    Exit For
                End If
            End If
        Next ilVehicle
       
        Close #hmProj
        mCloseProjFiles

        '11-03-05 client cannot handle a zipped file, send text only
        'initZIPCmdStruct
        'szFile = smExportName
        'zpcDZip.ZIPFile = smZipPath   'The ZIP file name"
        'zpcDZip.ItemList = szFile  'The file list to be added
        'zpcDZip.MajorStatusFlag = False  'Causes the major status event to trigger
        'zpcDZip.MinorStatusFlag = False  'Causes the minor status event to trigger
        'zpcDZip.BackgroundProcessFlag = True
        'zpcDZip.ActionDZ = ZIP_ADD   'ADD files to the ZIP file

        'ilRet = zpcDZip.ErrorCode       'return error code
        'If ilRet <> 0 Then
        '    MsgBox "Error in zipping. " & Str$(ilRet)
        '    Print #hmMsg, "Error in zipping db-tDailyProjection " & Str$(ilRet) & " "
        '    igBUerror = True
        'End If
        
        Screen.MousePointer = vbDefault

    'If ilRet = 0 Then           '11-03-05 file is no longer zipped.  true is successful, file is zipped
        'On Error GoTo mExportErrErr
        If StrComp(smExportName, smZipPath, 1) <> 0 Then       'if same path, no killing and no copying
            Kill smZipPath          'kill whats already in destination folder so new one can be copied

            FileCopy smExportName, smZipPath
        End If
        
        '11-03-05 client cannot handle a zipped file, send text only
        'lacInfo(0).Caption = "Export Successfully Completed-zipped in " & smZipPath
        If ilErrorStatus = 0 Then
            lacInfo(0).Caption = "Export Successfully Completed-saved in " & smZipPath
    
            gLogMsg "** Corporate Export Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "Corporate Export.txt", False
            
            '11-03-05 client cannot handle a zipped file, send text only
            'Print #hmMsg, "** Export has been zipped and saved in " & smZipPath
            gLogMsg "** Export has been saved in " & smZipPath, "Corporate Export.txt", False
        Else
            lacInfo(0).Caption = "Partial Export-saved in " & smZipPath
            gLogMsg "** Corporate Export partially completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "Corporate Export.txt", False
            gLogMsg "** Export has been saved in " & smZipPath, "Corporate Export.txt", False
        End If
    Else
        lacInfo(0).Caption = "Open Error: Export Failed"
        gLogMsg "** Export Open error : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "Corporate Export.txt", False
    End If
    'Else
    '    lacInfo(0).Caption = "Export Failed"
    '    Print #hmMsg, "** Export Failed **"
    'End If
    'lacInfo(1).Caption = "Export File: " & smExportName
    lacInfo(0).Visible = True
    'lacInfo(1).Visible = True
    'cmcCancel.Caption = "&Done"
    'cmcCancel.SetFocus
    ''cmcExport.Enabled = False
    Screen.MousePointer = vbDefault
    imExporting = False
    Erase tlChfAdvtExt
    
    If imAutoRun = 1 Then     '9-7-05 if auto run, task finished, unload and return to caller
        gUpdateTaskMonitor 2, "CE"
        lacStop_Click
    End If
    Exit Function
'mExportErrErr:
'    ilRet = Err.Number
'    Resume Next
                               'duplicate name, use next letter
'    ilRet = 1
'    Resume Next
End Function

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "CE"
End Sub

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    
    On Error Resume Next
    
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    igGGFlag = 1
    hmSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    imSafRecLen = Len(tmSaf)
    ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    ilField1 = Asc(tmSaf.sName)
    slStr = Mid$(tmSaf.sName, 2, 5)
    llDate = Val(slStr)
    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
    llNow = gDateValue(Format$(Now, "m/d/yy"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        igGGFlag = 0
    End If
    gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub
