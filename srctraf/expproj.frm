VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ExpProj 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4185
   ClientLeft      =   825
   ClientTop       =   2400
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4185
   ScaleWidth      =   7530
   Begin VB.TextBox edcEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1080
      Width           =   1200
   End
   Begin VB.TextBox edcStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   6
      Top             =   660
      Width           =   1200
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Vehicles"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "expproj.frx":0000
      Left            =   3000
      List            =   "expproj.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox edcContract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   10
      Top             =   1500
      Width           =   1200
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   840
      Top             =   3240
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   6840
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6450
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3270
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5835
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3270
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6105
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3135
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcTo 
      Appearance      =   0  'Flat
      Caption         =   "&Browse..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   15
      Top             =   2280
      Width           =   1485
   End
   Begin VB.PictureBox plcTo 
      Height          =   375
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   4245
      TabIndex        =   0
      Top             =   2160
      Width           =   4305
      Begin VB.TextBox edcTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   0
         TabIndex        =   14
         Top             =   30
         Width           =   4185
      End
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   3720
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   17
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label lacScreen 
      Appearance      =   0  'Flat
      Caption         =   "Corporate Export"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   2025
   End
   Begin DZACTXLibCtl.dzactxctrl zpcDZip 
      Left            =   1680
      OleObjectBlob   =   "expproj.frx":0004
      Top             =   3240
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   600
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label lacEnd 
      Appearance      =   0  'Flat
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      TabIndex        =   7
      Top             =   1140
      Width           =   1065
   End
   Begin VB.Label lacStart 
      Appearance      =   0  'Flat
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label lacContract 
      Appearance      =   0  'Flat
      Caption         =   "Contract #"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label lacSaveIn 
      Appearance      =   0  'Flat
      Caption         =   "Save In"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   810
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   3315
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   6255
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ExpProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''******************************************************************************************
''***** VB Compress Pro 6.11.32 generated this copy of expproj.frm on Wed 6/17/09 @ 12:56 PM
''***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Private Variables (Removed)                                                            *
''*  tmSdfSrchKey3                 tmSpot                                                  *
''*                                                                                        *
''* Public Procedures (Marked)                                                             *
''*  initZIPCmdStruct                                                                      *
''******************************************************************************************
'
'' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
'' Proprietary Software®, Do not copy
''
'' File Name: ExpProj.Frm
''
'' Release: 5.3
''
'' Description:
''   This file contains the Corporate Export (Projection & Avails) input screen
'Option Explicit
'Option Compare Text
'Dim hmMsg As Integer
'Dim hmProj As Integer
'Dim tmProjByDay() As PROJINFO
'Dim tmTradeCnts() As Long       'contract code of trade contracts (partial or full)
'
'Dim smExportName As String
'Dim smZipPath As String         'zipped path (base is from revenueexportpath in traffic.ini
'Dim imFirstActivate As Integer
'Dim lmCntrNo As Long    'for debugging purposes to filter a single contract
'Dim lmCntrCode As Long  'contract code so header doesnt have to be read to match all spots
'
'Dim imSetAll As Integer
'Dim imAllClicked As Integer
'
'Dim hmChf As Integer            'Contract header file handle
'Dim imChfRecLen As Integer        'CHF record length
'Dim tmChf As CHF
'Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
'
'Dim hmClf As Integer            'Contract line file handle
'Dim imClfRecLen As Integer        'CLF record length
'Dim tmClf As CLF
'Dim tmClfSrchKey As CLFKEY0
'
'Dim hmCff As Integer            'Contract flight file handle
'Dim imCffRecLen As Integer      'CFF record length
'Dim tmCff As CFF
'
'Dim hmVef As Integer            'Vehicle file handle
'Dim tmVef As VEF                'VEF record image
'Dim imVefRecLen As Integer        'VEF record length
'
'Dim hmVsf As Integer            'Vehicle file handle
'Dim tmVsf As VEF                'VSF record image
'Dim imVsfRecLen As Integer       'VSF record length
'
'Dim hmSdf As Integer            'Spot file handle
'Dim tmSdf As SDF                'Spot detail record image
'Dim imSdfRecLen As Integer        'SDF record length
'
'Dim hmSmf As Integer            'Spot MG file handle
'Dim tmSmf As SMF                'Spot MG record image
'Dim imSmfRecLen As Integer        'Spt MG record length
'
'Dim hmSsf As Integer
'Dim tmSsf As SSF                'Spot summary record image
'Dim imSsfRecLen As Integer        'SSF record length
'Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
'Dim tmSsfSrchKey2 As SSFKEY2 'SSF key record image
'Dim tmAvail As AVAILSS
'
'Dim hmLcf As Integer
'Dim tmLcf As LCF                'Log calendar  record image
'Dim imLcfRecLen As Integer        'Log calendar record length
'
'Dim imTerminate As Integer
'Dim imBypassFocus As Integer
'Dim imExporting As Integer
'Dim lmNowDate As Long
'Dim lmLastYearStartDate As Long    'earliest date to gather (jan 1 of previous year from todays date)
'Dim lmUserStartDate As Long         'user entered start date
'Dim lmUserEndDate As Long           'user entered end date
'Dim imAutoRun As Integer            '1 = auto run flag, 0 = manual
''*******************************************************
''*                                                     *
''*      Procedure Name:mParseCmmdLine                  *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Parse command line             *
''*                                                     *
''*******************************************************
'Private Sub mParseCmmdLine()
'    Dim slCommand As String
'    Dim slStr As String
'    Dim ilRet As Integer
'    Dim slTestSystem As String
'    Dim ilTestSystem As Integer
'
'    slCommand = sgCommandStr    'Command$
'    'If StrComp(slCommand, "Debug", 1) = 0 Then
'    '    igStdAloneMode = True 'Switch from/to stand alone mode
'    '    sgCallAppName = ""
'    '    slStr = "Guide"
'    '    ilTestSystem = False
'    '    imShowHelpMsg = False
'    'Else
'    '    igStdAloneMode = False  'Switch from/to stand alone mode
'        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
'        If Trim$(slStr) = "" Then
'            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
'            'End
'            imTerminate = True
'            Exit Sub
'        End If
'        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
'        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
'        If StrComp(slTestSystem, "Test", 1) = 0 Then
'            ilTestSystem = True
'        Else
'            ilTestSystem = False
'        End If
''        imShowHelpMsg = True
''        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
''        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
''            imShowHelpMsg = False
''        End If
'        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
'    'End If
'    'gInitStdAlone ExpPhnx, slStr, ilTestSystem
'    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
'    igCmmCallSource = Val(slStr)
'
'    ilRet = gParseItem(slCommand, 4, "\", slStr)
'    imAutoRun = Val(slStr)          'auto run flag 0= manual, 1 = auto
'    'If igStdAloneMode Then
'    '    igCmmCallSource = CALLNONE
'    'End If
'End Sub
'' **************************************************************************************
''
''  Procedure:  initZIPCmdStruct()
''
''  Purpose:  Set the ZIP control values
''
'' **************************************************************************************
'Sub initZIPCmdStruct() 'VBC NR
'  zpcDZip.ActionDZ = NO_ACTION 'VBC NR
'  zpcDZip.AddCommentFlag = False 'VBC NR
'  zpcDZip.AfterDateFlag = False 'VBC NR
'  zpcDZip.BackgroundProcessFlag = False 'VBC NR
'  zpcDZip.COMMENT = "" 'VBC NR
'  zpcDZip.CompressionFactor = 5 'VBC NR
'  zpcDZip.ConvertLFtoCRLFFlag = False 'VBC NR
'  zpcDZip.Date = "" 'VBC NR
'  zpcDZip.DeleteOriginalFlag = False 'VBC NR
'  zpcDZip.DiagnosticFlag = False 'VBC NR
'  zpcDZip.DontCompressTheseSuffixesFlag = False 'VBC NR
'  zpcDZip.DosifyFlag = False 'VBC NR
'  zpcDZip.EncryptCode = ""  'gCreatePassword 'VBC NR
'  zpcDZip.EncryptFlag = False   'True 'VBC NR
'  'zpcDZip.ExcludeFollowing = ""
'  'zpcDZip.ExcludeFollowingFlag = False
'  zpcDZip.FixFlag = False 'VBC NR
'  zpcDZip.FixHarderFlag = False 'VBC NR
'  zpcDZip.GrowExistingFlag = False 'VBC NR
'  zpcDZip.IncludeFollowing = "" 'VBC NR
'  zpcDZip.IncludeOnlyFollowingFlag = False 'VBC NR
'  zpcDZip.IncludeSysandHiddenFlag = False 'VBC NR
'  zpcDZip.IncludeVolumeFlag = False 'VBC NR
'  zpcDZip.ItemList = "" 'VBC NR
'  zpcDZip.MajorStatusFlag = True 'VBC NR
'  zpcDZip.MessageCallbackFlag = True 'VBC NR
'  zpcDZip.MinorStatusFlag = True 'VBC NR
'  zpcDZip.MultiVolumeControl = 0 'VBC NR
'
'  'Changed both of these to False from the default True
'  zpcDZip.NoDirectoryEntriesFlag = True 'VBC NR
'  zpcDZip.NoDirectoryNamesFlag = True 'VBC NR
'
'  zpcDZip.OldAsLatestFlag = False 'VBC NR
'  zpcDZip.PathForTempFlag = False 'VBC NR
'  zpcDZip.QuietFlag = False 'VBC NR
'  zpcDZip.RecurseFlag = False 'VBC NR
'  zpcDZip.StoreSuffixes = "" 'VBC NR
'  zpcDZip.TempPath = "" 'VBC NR
'  zpcDZip.ZIPFile = "" 'VBC NR
'
'  'Write out a log file in the windows sub directory
'  zpcDZip.ZipSubOptions = 256 'VBC NR
'
'  ' added for rev 3.00
'  zpcDZip.RenameCallbackFlag = False 'VBC NR
'  zpcDZip.ExtProgTitle = "" 'VBC NR
'  zpcDZip.ZIPString = "" 'VBC NR
'
'End Sub 'VBC NR
'
''*******************************************************
''*                                                     *
''*      Procedure Name:mGetCost                        *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Get Cost from line             *
''*                                                     *
''*******************************************************
'Private Function mGetCost(tlSdf As SDF, hlClf As Integer, hlCff As Integer, hlSmf As Integer, hlVef As Integer, hlVsf As Integer) As Long
'    Dim ilRet As Integer
'    Dim slPrice As String
'    mGetCost = 0
'    If tlSdf.sSpotType = "X" Then
'        Exit Function
'    End If
'    imClfRecLen = Len(tmClf)
'    tmClfSrchKey.lChfCode = tlSdf.lChfCode
'    tmClfSrchKey.iLine = tlSdf.iLineNo
'    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest Revision
'    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
'    ilRet = btrGetGreaterOrEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))
'        ilRet = btrGetNext(hlClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'    Loop
'    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) And ((tmClf.sSchStatus = "M") Or (tmClf.sSchStatus = "F")) Then
'        ilRet = gGetSpotPrice(tlSdf, tmClf, hlCff, hlSmf, hlVef, hlVsf, slPrice)
'        If InStr(slPrice, ".") > 0 Then
'            mGetCost = gStrDecToLong(slPrice, 2)
'        End If
'    End If
'End Function
'
''*******************************************************
''*                                                     *
''*      Procedure Name:mVehPop                         *
''*                                                     *
''*             Created:8/17/05       By:D. Hosaka      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Populate the selection combo   *
''*                      box for conventional           *
''*                      and selling vehicles           *
''*******************************************************
'Private Sub mVehPop()
'    Dim ilRet As Integer
'    ilRet = gPopUserVehicleBox(ExptGen, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)
'
'    If ilRet <> CP_MSG_NOPOPREQ Then
'        On Error GoTo mVehPopErr
'        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExpProj
'        On Error GoTo 0
'    End If
'
'    Exit Sub
'mVehPopErr:
'    On Error GoTo 0
'    imTerminate = True
'    Exit Sub
'End Sub
''********************************************************************************************
''           mCrProj- The Export gathers all spots for approx 24 months.
''                    Goes back to previous year starting jan 1 and
''                    continues to the future for as long as there are spots sched.
''                    A text file is created with inventory, avails, spot counts,
''                    revenue counts by hour within day and spot length by vehicle
''
''           Process by vehicle for all days, hours, spot lengths
''           11-15-05  selective contracts was never tested
''                     CCT wants to include all trades (partials & full) and
''                     always exclude fill spots
''********************************************************************************************
'Function mCrProj() As Integer
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilSpot                        tlChfAdvtExt                                            *
''******************************************************************************************
'
'Dim ilRet As Integer
'Dim ilError As Integer
'
'Dim ilIndex As Integer
'ReDim ilSpotLens(0 To 0) As Integer
'Dim ilLoop As Integer
'Dim ilUpper As Integer
'Dim llEarliestLcfDate As Long
'Dim llLatestLcfDate As Long
'Dim llDate As Long
'Dim ilDate(0 To 1) As Integer
'Dim llTime As Long
'Dim ilWhichHour As Integer
'Dim slDate As String
'Dim ilType As Integer
'ReDim tlSdf(0 To 0) As SDF
'
'    ilError = False             'assume everything is OK
'    ilIndex = gBinarySearchVpf(tmVef.iCode)
'    If ilIndex <> -1 Then
'        mSortLengths ilIndex, ilSpotLens()      'sort the spot lengths, to be used
'
'
'        If lmUserStartDate = 0 Then     'if no date entered, determine the earliest from the log calendar
'            llEarliestLcfDate = gGetEarliestLCFDate(hmLcf, "C", tmVef.iCode)
'            If llEarliestLcfDate > lmLastYearStartDate Then     'check that the start date of last year has log calendar in effect,
'                                                                'if not, use the earliest log calendar date
'                lmLastYearStartDate = llEarliestLcfDate
'            End If
'        Else                            'user entered a date
'            lmLastYearStartDate = lmUserStartDate
'        End If
'
'        If lmUserEndDate = 0 Then       'if no date entered, determine the latest date from the log calendar
'            llLatestLcfDate = gGetLatestLCFDate(hmLcf, "C", tmVef.iCode)     'determine how far in the future to obtain stats
'        Else
'            llLatestLcfDate = lmUserEndDate
'        End If
'
'        'show Processing vehicle & dates on caption screen & file
'        Print #hmMsg, "Processing " & Trim$(tmVef.sName) & " for " & Format(lmLastYearStartDate, "m/d/yy") & " - " & Format(llLatestLcfDate, "m/d/yy")
'        lacInfo(1).Caption = "Processing " & Trim$(tmVef.sName) & " for " & Format(lmLastYearStartDate, "m/d/yy") & " - " & Format(llLatestLcfDate, "m/d/yy")
'        lacInfo(1).Visible = True
'        ilType = 0
'        'use the ssf for inventory only
'        For llDate = lmLastYearStartDate To llLatestLcfDate
'
'            'Create arrays for each spot length for this day- all stats will be built in them
'            ilUpper = 0
'            ReDim tmProjByDay(0 To 0) As PROJINFO              'init for next vehicle
'            For ilLoop = 0 To UBound(ilSpotLens) - 1
'                tmProjByDay(ilLoop).iLen = ilSpotLens(ilLoop)
'                ilUpper = ilUpper + 1
'                ReDim Preserve tmProjByDay(0 To ilUpper) As PROJINFO
'            Next ilLoop
'
'            gPackDateLong llDate, ilDate(0), ilDate(1)
'            imSsfRecLen = Len(tmSsf)
'            If tmVef.sType <> "G" Then
'                tmSsfSrchKey.iType = 0 'slType-On Air
'                tmSsfSrchKey.iVefCode = tmVef.iCode
'                tmSsfSrchKey.iDate(0) = ilDate(0)
'                tmSsfSrchKey.iDate(1) = ilDate(1)
'                tmSsfSrchKey.iStartTime(0) = 0
'                tmSsfSrchKey.iStartTime(1) = 0
'                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
'            Else
'                tmSsfSrchKey2.iVefCode = tmVef.iCode
'                tmSsfSrchKey2.iDate(0) = ilDate(0)
'                tmSsfSrchKey2.iDate(1) = ilDate(1)
'                ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
'                ilType = tmSsf.iType
'            End If
'            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = tmVef.iCode) And (tmSsf.iDate(0) = ilDate(0)) And (tmSsf.iDate(1) = ilDate(1))
'                For ilLoop = 1 To tmSsf.iCount Step 1
'                   LSet tmAvail = tmSsf.tPAS(ilLoop)
'                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
'                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
'                        'process the inventory for the designated hour
'                        ilWhichHour = (llTime \ 3600) + 1
'                        mBuildInvByLength ilWhichHour, ilSpotLens()
'                    End If
'                Next ilLoop
'                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
'                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                If tmVef.sType = "G" Then
'                    ilType = tmSsf.iType
'                End If
'            Loop
'            slDate = Format$(llDate, "m/d/yy")
'            mBuildSpotStats slDate, tlSdf()    'obtain the spots for a single date
'            mWriteExportRec slDate              'create up to 24 records for the day (ignore hours without inventory)
'        Next llDate
'    Else                    'didnt find vehicle options
'        ilError = True
'    End If
'
'    mCrProj = ilError
'End Function
'
''
''
''
''           mCloseProjectionfiles - Close all applicable files for
''                       projection Export
''
'Sub mCloseProjFiles()
'Dim ilRet As Integer
'    ilRet = btrClose(hmVef)
'    ilRet = btrClose(hmClf)
'    ilRet = btrClose(hmCff)
'    ilRet = btrClose(hmSdf)
'    ilRet = btrClose(hmSsf)
'    ilRet = btrClose(hmLcf)
'    ilRet = btrClose(hmVsf)
'    ilRet = btrClose(hmSmf)
'    ilRet = btrClose(hmChf)
'
'    btrDestroy hmVef
'    btrDestroy hmClf
'    btrDestroy hmCff
'    btrDestroy hmSdf
'    btrDestroy hmSsf
'    btrDestroy hmLcf
'    btrDestroy hmVsf
'    btrDestroy hmSmf
'    btrDestroy hmChf
'End Sub
'
''
''
''           mOpenProjectionFiles - open files applicable to Projection Export
''                           8-17-05
''
''
'Function mOpenProjFiles() As Integer
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilTemp                        slStamp                                                 *
''******************************************************************************************
'
'Dim ilRet As Integer
'Dim ilError As Integer
'
'    ilError = False
'
'    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen VEF)", ExpProj
'    On Error GoTo 0
'    imVefRecLen = Len(tmVef)
'
'    hmChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen CHF)", ExpProj
'    On Error GoTo 0
'    imChfRecLen = Len(tmChf)
'
'    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen CLF)", ExpProj
'    On Error GoTo 0
'    imClfRecLen = Len(tmClf)
'
'    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen CFF)", ExpProj
'    On Error GoTo 0
'    imCffRecLen = Len(tmCff)
'
'    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen SDF)", ExpProj
'    On Error GoTo 0
'    imSdfRecLen = Len(tmSdf)
'
'    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen SSF)", ExpProj
'    On Error GoTo 0
'    imSsfRecLen = Len(tmSsf)
'
'    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen LCF)", ExpProj
'    On Error GoTo 0
'    imLcfRecLen = Len(tmLcf)
'
'    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen SMF)", ExpProj
'    On Error GoTo 0
'    imSmfRecLen = Len(tmSmf)
'
'    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenProjFilesErr
'    gBtrvErrorMsg ilRet, "gOpenProjFiles (btrOpen VSF)", ExpProj
'    On Error GoTo 0
'    imVsfRecLen = Len(tmVsf)
'
'
'    mOpenProjFiles = ilError
'    Exit Function
'
'mOpenProjFilesErr:
'    ilError = True
'    Return
'End Function
'
''
''
''           mWriteExportRec - gather all the information for a day, hour, and spot length and write
''           a record to the export .txt file
''
''           <input> tlProjInfo - structure containing all the info required to previous years data
''                               starting jan1, and going thru the last spot scheduled for the vehicle.
''                               Create records by hour within date by vehicle and spot length.
''           Return - true if error, otherwise false
'Private Function mWriteExportRec(slDate As String) As Integer
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilRemainder                   slStr                                                   *
''******************************************************************************************
'
'Dim ilLoop As Integer
'Dim slRecord As String
'Dim ilIndex As Integer
'Dim slVehicle As String
'Dim ilError As Integer
'Dim ilRecdLen As Integer
'Dim slMonth As String
'Dim slYear As String
'Dim slAmt As String
'Dim llAvgRate As Long
'Dim ilWhichHour As Integer
'Dim ilRet As Integer
'Dim ilDay As Integer
'Dim slDay As String
'Dim slDayDescr As String * 3
'Dim slDayOfWeek As String * 21
'
'    slDayOfWeek = "MONTUEWEDTHUFRISATSUN"
'    ilError = False
'
'    'format the month info for a contract/vehicle
'    slVehicle = ""
'    ilIndex = gBinarySearchVef(tmVef.iCode)
'    If ilIndex <> -1 Then
'        slVehicle = Trim$(tmVef.sName)
'
'        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
'        'obtain the day of the week
'        ilDay = gWeekDayStr(slDate)
'        slDayDescr = Mid$(slDayOfWeek, (ilDay * 3) + 1, 3)
'        For ilLoop = LBound(tmProjByDay) To UBound(tmProjByDay) - 1
'            For ilWhichHour = 1 To 24
'                'ignore record if no inventory defined
'                If tmProjByDay(ilLoop).iSchedUnits(ilWhichHour) <> 0 Or tmProjByDay(ilLoop).iInventory(ilWhichHour) <> 0 Then        'need at least 1 spot sch or inv to create record
'                    slRecord = Trim$(tmVef.sName) & Chr$(9)              'vehicle name & tab
'                    'slRecord = slRecord & Val(tmVef.sCodeStn) & Chr$(9)       'this is a string field, user needs an integer so it must
'                                                                     'this field should be a number input
'                                                                     'dont need this as long as they have call letters and band
'
'                    slRecord = slRecord & slYear & "-" & slMonth & "-" & Trim$(slDay) & " " & Trim$(str(ilWhichHour - 1)) & ":00:00" & Chr$(9)   'date ,  hour (military) , tab
'
'                    slRecord = slRecord & slDayDescr & Chr$(9)          'day of week descr, tab
'                    slRecord = slRecord & Trim$(str(tmProjByDay(ilLoop).iLen)) & Chr$(9)        'spot length, tab
'
'                    slRecord = slRecord & "0" & Chr$(9)                 'BTA (always 0) , tab
'                    slRecord = slRecord & Trim$(str(tmProjByDay(ilLoop).iInventory(ilWhichHour))) & Chr$(9)  'inventory , tab
'
'                    slAmt = gLongToStrDec(tmProjByDay(ilLoop).lMinRate(ilWhichHour), 2)
'                    slRecord = slRecord & slAmt & Chr$(9)           'min $ value , tab
'
'                    slAmt = gLongToStrDec(tmProjByDay(ilLoop).lMaxRate(ilWhichHour), 2)
'                    slRecord = slRecord & slAmt & Chr$(9)           'max $ value , tab
'
'                    If tmProjByDay(ilLoop).iSchedUnits(ilWhichHour) <> 0 Then
'                        llAvgRate = tmProjByDay(ilLoop).lSchedRev(ilWhichHour) / tmProjByDay(ilLoop).iSchedUnits(ilWhichHour) 'determine avg rate (sched revenue/sched units)
'                        slAmt = gLongToStrDec(llAvgRate, 2)
'                    Else
'                        slAmt = ".00"
'                    End If
'                    slRecord = slRecord & slAmt & Chr$(9)               'avg unit rate, tab
'
'                    slAmt = gLongToStrDec(tmProjByDay(ilLoop).lSchedRev(ilWhichHour), 2)
'                    slRecord = slRecord & slAmt & Chr$(9)           'schedule revenue , tab
'
'                    slAmt = gLongToStrDec(tmProjByDay(ilLoop).lMissedRev(ilWhichHour), 2)
'                    slRecord = slRecord & slAmt & Chr$(9)           'missed (pooled) revenue , tab
'
'                    slAmt = Trim$(str(tmProjByDay(ilLoop).iSchedUnits(ilWhichHour)))
'                    slRecord = slRecord & slAmt & Chr$(9)           'scheduled units , tab
'
'
'                    slAmt = Trim$(str(tmProjByDay(ilLoop).iNCUnits(ilWhichHour)))
'                    slRecord = slRecord & slAmt & Chr$(9)           'zero rate units , tab
'
'                    slAmt = Trim$(str(tmProjByDay(ilLoop).iMissedUnits(ilWhichHour)))
'                    slRecord = slRecord & slAmt & Chr$(9)           'missed (pooled) units , tab
'
'                    llAvgRate = tmProjByDay(ilLoop).iInventory(ilWhichHour) - tmProjByDay(ilLoop).iSchedUnits(ilWhichHour)
'                    slAmt = Trim$(str(llAvgRate))
'                    slRecord = slRecord & slAmt & Chr$(9)               'avails, tab
'
'                    slRecord = slRecord & "0" & Chr$(9)                 'unused time (always 0), tab
'                    slRecord = slRecord & Chr$(13) & Chr$(10)       'carriage return, line Chr$(9)
'
'                    ilRecdLen = Len(Trim(slRecord))
'                    On Error GoTo mWriteExportRecErr
'                    Print #hmProj, Left(slRecord, ilRecdLen)
'                    On Error GoTo 0
'                    If ilRet <> 0 Then
'                        imExporting = False
'                        Print #hmMsg, "** Terminated **"
'                        Close #hmMsg
'                        Close #hmProj
'                        Screen.MousePointer = vbDefault
'                        MsgBox "Error writing to " & sgDBPath & "Messages\" & "ExpProj.Txt" & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
'                        cmcCancel.SetFocus
'                        Exit Function
'                    End If
'                End If                          'zero inventory
'            Next ilWhichHour
'        Next ilLoop
'
'    End If
'
'    mWriteExportRec = ilError
'    Exit Function
'
'mWriteExportRecErr:
'    ilError = True
'    Resume Next
'
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mOpenMsgFile                    *
''*                                                     *
''*             Created:5/18/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Open error message file         *
''*                                                     *
''*******************************************************
'Private Function mOpenMsgFile()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  slNTR                         slCntr                                                  *
''******************************************************************************************
'
'    Dim slToFile As String
'    Dim slDateTime As String
'    Dim slFileDate As String
'    Dim ilRet As Integer
'
'    ilRet = 0
'    On Error GoTo mOpenMsgFileErr:
'    slToFile = sgDBPath & "\Messages\" & "ExpProj.Txt"
'    slDateTime = FileDateTime(slToFile)
'    If ilRet = 0 Then
'        slFileDate = Format$(slDateTime, "m/d/yy")
'        If gDateValue(slFileDate) = lmNowDate Then  'Append
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Append As hmMsg
'            If ilRet <> 0 Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        Else
'            Kill slToFile
'            On Error GoTo 0
'            ilRet = 0
'            On Error GoTo mOpenMsgFileErr:
'            hmMsg = FreeFile
'            Open slToFile For Output As hmMsg
'            If ilRet <> 0 Then
'                Screen.MousePointer = vbDefault
'                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'                mOpenMsgFile = False
'                Exit Function
'            End If
'        End If
'    Else
'        On Error GoTo 0
'        ilRet = 0
'        On Error GoTo mOpenMsgFileErr:
'        hmMsg = FreeFile
'        Open slToFile For Output As hmMsg
'        If ilRet <> 0 Then
'            Screen.MousePointer = vbDefault
'            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'            mOpenMsgFile = False
'            Exit Function
'        End If
'    End If
'    On Error GoTo 0
'
'    Print #hmMsg, "** Corporate Export " & ": " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
'    'Print #hmMsg, ""
'
'    mOpenMsgFile = True
'    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
'End Function
'
'Private Sub ckcAll_Click()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  llValue                                                                               *
''******************************************************************************************
'
'Dim Value As Integer
'Dim llRg As Long
'Dim ilValue As Integer
'Dim llRet As Long
'
'    If lbcVehicle.ListCount <= 0 Then
'        Exit Sub
'    End If
'
'    Value = False
'    If ckcAll.Value = vbChecked Then
'        Value = True
'    End If
'
'    ilValue = Value
'    If imSetAll Then
'        imAllClicked = True
'        llRg = CLng(lbcVehicle.ListCount - 1) * &H10000 Or 0
'        llRet = SendMessageByNum(lbcVehicle.hwnd, LB_SELITEMRANGE, ilValue, llRg)
'        imAllClicked = False
'    End If
'    mSetCommands
'End Sub
'
'Private Sub cmcCancel_Click()
'    If imExporting Then
'        imTerminate = True
'        Exit Sub
'    End If
'    mTerminate
'End Sub
''
''
''                   Create Projection file (tab deliminted) that contains inventory from SSF
''                   and spot statistics from SDF.  Each unique record is broken out by date,
''                   spot length and hour.  Inventory is determined by the max lengths allowed.
''                   For example:  2/60 is treated as 1 60" unit, rather than 2 30s.  If a 30 is
''                   booked, it is considered oversold by 1 30", and open to 1 60" spot.
''                   Revenue is gathered as well as missed spots.
''                   PSAS and Promos are ignored, as well as trade contracts.
''
'Private Sub cmcExport_Click()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilLoop                        ilYear                        slStart                   *
''*  slName                        ilAnswer                      szFile                    *
''*                                                                                        *
''******************************************************************************************
'
'    Dim ilRet As Integer
'    Dim slStr As String
'    Dim slDateTime As String
'    Dim ilVehicle As Integer
'    Dim slNameCode As String
'    Dim slCode As String
'    Dim ilVefCode As Integer
'    Dim ilError As Integer
'    Dim ilVefIndex As Integer
'    Dim slStartDate As String
'    Dim slEndDate As String
'    Dim slCntrType As String
'    Dim slCntrStatus As String
'    Dim ilHOState As Integer
'    Dim ilChf As Integer
'    Dim ilUpper As Integer
'    Dim slYear As String
'    Dim slMonth As String
'    Dim slDay As String
'    ReDim tmTradeCnts(0 To 0) As Long
'
'
'    lacInfo(0).Visible = False
'
'    If imExporting Then
'        Exit Sub
'    End If
'
'
'    lmCntrNo = 0                'ths is for debugging on a single contract
'    slStr = ExpProj!edcContract
'    If slStr <> "" Then
'        lmCntrNo = Val(slStr)
'    End If
'
'    'smExportFile contains the name to use which has been moved to edcTo.Text
'    smExportName = Trim$(edcTo.Text)
'    If Len(smExportName) = 0 Then
'        Beep
'        edcTo.SetFocus
'        Exit Sub
'    End If
'
'    If (InStr(smExportName, ":") = 0) And (Left$(smExportName, 2) <> "\\") Then     'test for absence of colon and not using \\
'        smExportName = sgExportPath & smExportName
'    End If
'
'    ilRet = 0
'    On Error GoTo cmcExportErr:
'    slDateTime = FileDateTime(smExportName)
'    If ilRet = 0 Then
'        'always kill the duplicate file name
'
'        'file already exists, continue?
'        'ilAnswer = MsgBox("Filename already exists, overwrite?", vbYesNo + vbApplicationModal, "Save In")
'        'If ilAnswer = vbYes Then
'            Kill smExportName
'        'Else
'        '    Exit Sub
'        'End If
'    End If
'
'    If Not mOpenMsgFile() Then          'open message file
'         cmcCancel.SetFocus
'         Exit Sub
'    End If
'    On Error GoTo 0
'    ilRet = 0
'    On Error GoTo cmcExportErr:
'    hmProj = FreeFile
'    Open smExportName For Output As hmProj
'    If ilRet <> 0 Then
'        Print #hmMsg, "** Terminated **"
'        Close #hmMsg
'        Close #hmProj
'        imExporting = False
'        Screen.MousePointer = vbDefault
'        MsgBox "Open Error #" & str$(Err.Numner) & smExportName, vbOKOnly, "Open Error"
'        Exit Sub
'    End If
'    Print #hmMsg, "** Storing Output into " & smExportName & " **"
'
'    Screen.MousePointer = vbHourglass
'    imExporting = True
'
'    If mOpenProjFiles() = 0 Then
'        slStartDate = Format$(lmLastYearStartDate, "m/d/yy")
'        slStr = Format$(lmNowDate, "m/d/yy")
'        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
'        slEndDate = "12/31/" & Trim$(slYear)
'
'        'test for valid single contract if entered
'        If lmCntrNo <> 0 Then
'            tmChfSrchKey1.lCntrNo = lmCntrNo
'            tmChfSrchKey1.iCntRevNo = 32000
'            tmChfSrchKey1.iPropVer = 32000
'            ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'            If ilRet = BTRV_ERR_END_OF_FILE Or tmChf.lCntrNo <> lmCntrNo Then
'                'exit, no contract found
'                lacInfo(0).Caption = "Contract #" & Trim$(ExpProj!edcContract) & " does not exist, application aborted"
'                Print #hmMsg, "Contract #" & Trim$(ExpProj!edcContract) & " does not exist, application aborted"
'                Close #hmProj
'                mCloseProjFiles
'
'                lacInfo(0).Visible = True
'                Close #hmMsg
'                cmcCancel.Caption = "&Done"
'                cmcCancel.SetFocus
'                Screen.MousePointer = vbDefault
'                imExporting = False
'
'                If imAutoRun Then       '9-7-05 if auto run, task finished, unload and return to caller
'                    cmcCancel_Click
'                End If
'
'                Exit Sub
'            Else
'                lmCntrCode = tmChf.lCode
'            End If
'        End If
'
'        'build array of trade contrcts so the contract header doesnt have to be reread for each spot
'        slCntrStatus = "HOGN"                 'statuses: hold, order, unsch hold, uns order
'        slCntrType = "CVTRQ"         'all types: PI, DR, etc.  except PSA(p) and Promo(m)
'        ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)
'
'        ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
'        ilRet = gObtainCntrForDate(ExpProj, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
'        ilUpper = 0
'        For ilChf = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1
'
'            '11-14-05 CCT wants to always include all trades
'            'If tlChfAdvtExt(ilChf).iPctTrade = 100 Then         'only 100% trades excluded
'            '    tmTradeCnts(ilUpper) = tlChfAdvtExt(ilChf).lCode        'save the contr codes to match the spot records (sdfchfcode)
'            '    ilUpper = ilUpper + 1
'            '    ReDim Preserve tmTradeCnts(0 To ilUpper) As Long
'            'End If
'        Next ilChf
'
'        'get user entered dates for testing purposes
'        If edcStart.Text = "" Then
'            lmUserStartDate = 0
'        Else
'            lmUserStartDate = gDateValue(edcStart.Text)
'        End If
'        If edcEnd.Text = "" Then
'            lmUserEndDate = 0
'        Else
'            lmUserEndDate = gDateValue(edcEnd.Text)
'        End If
'
'        For ilVehicle = 0 To lbcVehicle.ListCount - 1
'            If lbcVehicle.Selected(ilVehicle) = True Then
'                slNameCode = tgUserVehicle(ilVehicle).sKey
'                ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                ilVefCode = Val(slCode)
'                ilVefIndex = gBinarySearchVef(ilVefCode)
'                tmVef = tgMVef(ilVefIndex)
'                ilError = mCrProj()
'                If ilError Then     'returned with error?
'                    lacInfo(0).Caption = Trim$(tmVef.sName) & " Options table missing"
'                    Print #hmMsg, " Export error - " & Trim$(tmVef.sName) & " options table missing"
'                End If
'            End If
'        Next ilVehicle
'
'        Close #hmProj
'        mCloseProjFiles
'
'        '11-03-05 client cannot handle a zipped file, send text only
'        'initZIPCmdStruct
'        'szFile = smExportName
'        'zpcDZip.ZIPFile = smZipPath   'The ZIP file name"
'        'zpcDZip.ItemList = szFile  'The file list to be added
'        'zpcDZip.MajorStatusFlag = False  'Causes the major status event to trigger
'        'zpcDZip.MinorStatusFlag = False  'Causes the minor status event to trigger
'        'zpcDZip.BackgroundProcessFlag = True
'        'zpcDZip.ActionDZ = ZIP_ADD   'ADD files to the ZIP file
'
'        'ilRet = zpcDZip.ErrorCode       'return error code
'        'If ilRet <> 0 Then
'        '    MsgBox "Error in zipping. " & Str$(ilRet)
'        '    Print #hmMsg, "Error in zipping db-tDailyProjection " & Str$(ilRet) & " "
'        '    igBUerror = True
'        'End If
'
'        Screen.MousePointer = vbDefault
'    Else
'        lacInfo(0).Caption = "Open Error: Export Failed"
'        Print #hmMsg, "** Export Open error : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
'    End If
'
'    'If ilRet = 0 Then           '11-03-05 file is no longer zipped.  true is successful, file is zipped
'        On Error GoTo cmcExportErr
'        If StrComp(smExportName, smZipPath, 1) <> 0 Then       'if same path, no killing and no copying
'            Kill smZipPath          'kill whats already in destination folder so new one can be copied
'
'            FileCopy smExportName, smZipPath
'        End If
'
'        '11-03-05 client cannot handle a zipped file, send text only
'        'lacInfo(0).Caption = "Export Successfully Completed-zipped in " & smZipPath
'        lacInfo(0).Caption = "Export Successfully Completed-saved in " & smZipPath & " and " & smExportName
'
'        Print #hmMsg, "** Corporate Export Successfully completed : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
'
'        '11-03-05 client cannot handle a zipped file, send text only
'        'Print #hmMsg, "** Export has been zipped and saved in " & smZipPath
'        Print #hmMsg, "** Export has been saved in " & smZipPath
'
'    'Else
'    '    lacInfo(0).Caption = "Export Failed"
'    '    Print #hmMsg, "** Export Failed **"
'    'End If
'    'lacInfo(1).Caption = "Export File: " & smExportName
'    lacInfo(0).Visible = True
'    'lacInfo(1).Visible = True
'    Close #hmMsg
'    cmcCancel.Caption = "&Done"
'    cmcCancel.SetFocus
'    'cmcExport.Enabled = False
'    Screen.MousePointer = vbDefault
'    imExporting = False
'    Erase tlChfAdvtExt
'    Erase tmTradeCnts
'    Erase tmProjByDay
'
'    If imAutoRun Then       '9-7-05 if auto run, task finished, unload and return to caller
'        cmcCancel_Click
'    End If
'    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next
'                               'duplicate name, use next letter
'    ilRet = 1
'    Resume Next
'End Sub
'Private Sub cmcTo_Click()
'    CMDialogBox.DialogTitle = "Export To File"
'    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
'    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
'    CMDialogBox.DefaultExt = ".Csv"
'    CMDialogBox.flags = cdlOFNCreatePrompt
'    CMDialogBox.Action = 1 'Open dialog
'    edcTo.Text = CMDialogBox.fileName
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'        ChDir sgCurDir
'    End If
'    If edcTo.Text = "" Then
'        edcTo.Text = smExportName
'    End If
'End Sub
'Private Sub edcContract_GotFocus()
'    gCtrlGotFocus edcContract
'End Sub
'
'Private Sub edcEnd_GotFocus()
'    gCtrlGotFocus edcEnd
'End Sub
'
'Private Sub edcLinkDestHelpMsg_Change()
'    igParentRestarted = True
'End Sub
'
'
'
'Private Sub edcStart_GotFocus()
'    gCtrlGotFocus edcStart
'End Sub
'
'Private Sub edcTo_Change()
'    mSetCommands
'End Sub
'
'Private Sub Form_Activate()
'    If Not imFirstActivate Then
'        DoEvents    'Process events so pending keys are not sent to this
'                    'form when keypreview turn on
'        Me.KeyPreview = True
'        Exit Sub
'    End If
'    imFirstActivate = False
'    DoEvents    'Process events so pending keys are not sent to this
'    Me.KeyPreview = True
'    Me.Refresh
'
'End Sub
'
'Private Sub Form_Deactivate()
'    Me.KeyPreview = False
'End Sub
'
'Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
'    sgDoneMsg = CmdStr
'    igChildDone = True
'    Cancel = 0
'End Sub
'Private Sub Form_Load()
'    mInit
'    If imTerminate Then
'        cmcCancel_Click
'    End If
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
''Rm**    ilRet = btrReset(hgHlf)
''Rm**    btrDestroy hgHlf
'    'btrStopAppl
'    'End
'End Sub
'Private Sub imcHelp_Click()
'
'    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
'    'Traffic!cdcSetup.Action = 6
'End Sub
'
''*******************************************************
''*                                                     *
''*      Procedure Name:mInit                           *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Initialize modular             *
''*                                                     *
''*******************************************************
'Private Sub mInit()
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilMonth                       ilYear                        slEndDate                 *
''*                                                                                        *
''******************************************************************************************
'
''
''   mInit
''   Where:
''
'    Dim ilRet As Integer
'    Dim slTodayDate As String
'    Dim slLastYearDate As String
'    Dim slDay As String
'    Dim slMonth As String
'    Dim slYear As String
'    Dim slNameCode As String
'    Dim slCode As String
'    Dim ilVpf As Integer
'    Dim ilVefCode As Integer
'    Dim ilVehicle As Integer
'
'    mParseCmmdLine
'
'    imTerminate = False
'    imFirstActivate = True
'    'mParseCmmdLine
'    Screen.MousePointer = vbHourglass
'    imExporting = False
'    imBypassFocus = False
'    imSetAll = True
'    imAllClicked = False
'    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
'
'    gCenterStdAlone ExpProj
'
'    ilRet = gObtainVef() 'Build into tgMVef
'    If ilRet = False Then
'        imTerminate = True
'    End If
'    mVehPop
'
'    For ilVehicle = 0 To lbcVehicle.ListCount - 1
'        slNameCode = tgUserVehicle(ilVehicle).sKey
'        ilRet = gParseItem(slNameCode, 2, "\", slCode)
'        ilVefCode = Val(slCode)
'        ilVpf = gBinarySearchVpf(ilVefCode)
'        If ilVpf <> -1 Then
'            If (tgVpf(ilVpf).sExpHiCorp <> "N") Then     'this vehicle flagged to export
'                lbcVehicle.Selected(ilVehicle) = True
'            End If
'        End If
'     Next ilVehicle
'
'    slTodayDate = Format$(lmNowDate, "m/d/yy")
'    gObtainYearMonthDayStr slTodayDate, True, slYear, slMonth, slDay
'    slLastYearDate = "1/1/" & Trim$(str$(Val(slYear) - 1))
'    edcStart.Text = slLastYearDate
'    lmLastYearStartDate = gDateValue(slLastYearDate)             'determine start of last years date
'
'    Screen.MousePointer = vbDefault
'
'    tmcClick.Interval = 2000    '2 seconds
'    tmcClick.Enabled = True
'    Exit Sub
'
'End Sub
'
'
'
''*******************************************************
''*                                                     *
''*      Procedure Name:mTerminate                      *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: terminate form                 *
''*                                                     *
''*******************************************************
'Private Sub mTerminate()
''
''   mTerminate
''   Where:
''
'    Dim ilRet As Integer
'
'    ilRet = btrClose(hmVef)
'    ilRet = btrClose(hmChf)
'    ilRet = btrClose(hmClf)
'    ilRet = btrClose(hmCff)
'    ilRet = btrClose(hmSsf)
'    ilRet = btrClose(hmSdf)
'    ilRet = btrClose(hmLcf)
'    ilRet = btrClose(hmVsf)
'    ilRet = btrClose(hmSmf)
'
'    btrDestroy hmVef
'    btrDestroy hmChf
'    btrDestroy hmClf
'    btrDestroy hmCff
'    btrDestroy hmSsf
'    btrDestroy hmSdf
'    btrDestroy hmLcf
'    btrDestroy hmSmf
'    btrDestroy hmVsf
'    Screen.MousePointer = vbDefault
'    'igParentRestarted = False
'    'If Not igStdAloneMode Then
'    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
'    '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
'    '    Else
'    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
'    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
'    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
'    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
'    '        edcLinkDestHelpMsg.LinkExecute "Done"
'    '    End If
'    '    Do While Not igParentRestarted
'    '        DoEvents
'    '    Loop
'    'End If
'    Screen.MousePointer = vbDefault
'    igManUnload = YES
'    Unload ExpProj
'    Set ExpProj = Nothing   'Remve data segment
'    igManUnload = NO
'End Sub
'
'Private Sub lbcVehicle_Click()
'    If Not imAllClicked Then
'        imSetAll = False
'        ckcAll.Value = vbUnchecked  '9-12-02 False
'        imSetAll = True
'    End If
'    mSetCommands
'End Sub
'
'
'Private Sub tmcClick_Timer()
'Dim slRepeat As String
'Dim ilRet As Integer
'Dim slDateTime As String
'Dim slDefaultFileName As String
'
'    tmcClick.Enabled = False
'    'Determine name of export (.txt file)
'    'for now, do not create unique names that append a new letter of the alphabet.
'    'Always ask to overwrite
'    slRepeat = ""
'    slDefaultFileName = "dbo-tDailyProjection" & Trim$(slRepeat)
'    'Do
'        ilRet = 0
'        On Error GoTo cmcExportDupNameErr:
'        smExportName = sgExportPath & slDefaultFileName & ".txt"
'        If sgRevenueExportPath = "" Then
'            smZipPath = sgExportPath & slDefaultFileName & ".txt"           '11-3-05 append extension
'        Else
'            smZipPath = sgRevenueExportPath & slDefaultFileName & ".txt"    '11-3-05 append extension
'        End If
'        slDateTime = FileDateTime(smExportName)
'        'If ilRet = 0 Then       'fell thru , there was a filename that existed with same name. Increment the letter
'        '    If slRepeat = "" Then       'first time doesnt have the alpha appended for conseutive runs
'        '        slRepeat = "A"
'        '    Else
'        '        slRepeat = Chr(Asc(slRepeat) + 1)
'        '    End If
'        'End If
'    'Loop While ilRet = 0
'    edcTo.Text = smExportName
'    edcTo.Visible = True
'
'    If imAutoRun Then           '9-07-05 if auto run, activate the task
'        '9-30-05 changed to use field in vpf (vpf.sExpHiCorp to indicate if  vehicle should be included)
'        'ckcAll.Value = vbChecked    'assume all vehicles
'        cmcExport.Enabled = True
'        cmcExport_Click
'    End If
'    Exit Sub
'cmcExportDupNameErr:
'    ilRet = 1
'    Resume Next
'End Sub
'
'Public Sub mSetCommands()
'
'    If lbcVehicle.SelCount > 0 And edcTo.Text <> "" Then
'        cmcExport.Enabled = True
'    Else
'        cmcExport.Enabled = False
'    End If
'
'End Sub
'
''
''               msortLengths - sort the spots lengths stored with
''               vehicle options in descending order.  Used to determine
''               inventory in breaking them out by lengths.  Filter out
''               the items unused (zero lengths)
'Private Sub mSortLengths(ilIndex As Integer, ilSpotLens() As Integer)
'Dim ilUpper As Integer
'Dim ilLoop As Integer
'Dim ilEventLoop As Integer
'Dim ilCounter As Integer
'Dim ilTemp As Integer
'
'    ilUpper = 0
'    For ilLoop = 0 To 9
'        If tgVpf(ilIndex).iSLen(ilLoop) > 0 Then
'            ilSpotLens(ilUpper) = tgVpf(ilIndex).iSLen(ilLoop)
'            ilUpper = ilUpper + 1
'            ReDim Preserve ilSpotLens(0 To ilUpper)
'        End If
'    Next ilLoop
'
'    For ilEventLoop = 0 To ilUpper - 1
'        For ilCounter = 0 To ilUpper - 2
'
'            If ilSpotLens(ilCounter) < ilSpotLens(ilCounter + 1) Then
'                ilTemp = ilSpotLens(ilCounter)
'                ilSpotLens(ilCounter) = ilSpotLens(ilCounter + 1)
'                ilSpotLens(ilCounter + 1) = ilTemp
'            End If
'        Next ilCounter
'    Next ilEventLoop
'    Exit Sub
'End Sub
''
''               mBuildInvBylength - accumulate the amount of inventory
''               for the avail.  Use the spot length table to determine
''               which lengths to use for inventory.  Inv will be
''               calculated using the highest spot length down to
''               the lowest.  For example:  Vehicle spots lengths: 60 30 15 10
''               A 2M45S avail would result in 2 60 units,1 30 unit, 1 15 unit
''               <input> ilWhichHour - hour of day to build stats (1-24)
''                       ilSpotLens() - array of descending spot lengths from veh options
''               8-18-05
'Private Sub mBuildInvByLength(ilWhichHour As Integer, ilSpotLens() As Integer)
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilAvailLen                                                                            *
''******************************************************************************************
'
'Dim ilRemaining As Integer
'Dim ilLoop As Integer
'Dim ilUnits As Integer
'Dim ilStats As Integer
'
'    ilRemaining = tmAvail.iLen
'    For ilLoop = 0 To UBound(ilSpotLens) - 1
'        ilUnits = ilRemaining \ ilSpotLens(ilLoop)
'        ilRemaining = ilRemaining - (ilUnits * ilSpotLens(ilLoop))
'        'udpate the inventory statistics entry for the matching spot length
'        For ilStats = LBound(tmProjByDay) To UBound(tmProjByDay) - 1
'            If tmProjByDay(ilStats).iLen = ilSpotLens(ilLoop) Then
'                tmProjByDay(ilStats).iInventory(ilWhichHour) = tmProjByDay(ilStats).iInventory(ilWhichHour) + ilUnits
'            End If
'        Next ilStats
'    Next ilLoop
'    Exit Sub
'End Sub
''
''           mBuildspotStats - determine stats for revenue, spots sched/missed & availability
''           Read the spots by vehicle, date, time & sch status for one day at a time
''           Arrays are maintained by spot length.  Within that spot length array are
''           24 hour buckets representing revenue, inventory, sold, N/c spots, etc.
''           Ignore PSA, Promo & 100% trade contracts
''           11-14-05 CCT wants to include all trades (partials & full trade), and exlude fills
'Private Sub mBuildSpotStats(slDate As String, tlSdf() As SDF)
'Dim ilStats As Integer
'Dim llAmt As Long
'Dim ilSpots As Integer
'Dim ilRet As Integer
'Dim ilWhichHour As Integer
'Dim llTime As Long
'Dim ilTrade As Integer
'Dim ilSpotOK As Integer
'
'    ilRet = gGetSpotsbyVefDate(hmSdf, tmVef.iCode, slDate, slDate, tlSdf())
'    For ilSpots = LBound(tlSdf) To UBound(tlSdf) - 1
'        tmSdf = tlSdf(ilSpots)
'        'ignore psa, promo, fills (11-14-05) and hidden spots
'        If (tmSdf.sSpotType <> "S" And tmSdf.sSpotType <> "M" And tmSdf.sSpotType <> "X" And tmSdf.sSchStatus <> "H") And (lmCntrCode = 0 Or lmCntrCode = tmSdf.lChfCode) Then        'exclude PSA/Promos, fills and hidden spots
'            gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
'            'process the scheduled spts for the designated hour
'            ilWhichHour = (llTime \ 3600) + 1
'            For ilStats = LBound(tmProjByDay) To UBound(tmProjByDay) - 1        'find matching spot length entry
'                If tmProjByDay(ilStats).iLen = tmSdf.iLen Then
'                    ilSpotOK = True
'                    '11-14-05 trades are always included; this array is never created
'                    For ilTrade = 0 To UBound(tmTradeCnts) - 1
'                        If tmSdf.lChfCode = tmTradeCnts(ilTrade) Then   'if matching contract header code, this is a trade and
'                                                                        'should be ignored
'                            ilSpotOK = False
'                            Exit For
'                        End If
'                    Next ilTrade
'                    If ilSpotOK Then
'                        If tmSdf.sSpotType = "X" Or tmSdf.sSpotType = "O" Or tmSdf.sSpotType = "C" Then 'no rates apply for fill spots, or bb open/close spots
'                            llAmt = 0
'                        Else
'                            llAmt = mGetCost(tmSdf, hmClf, hmCff, hmSmf, hmVef, hmVsf)
'                        End If
'
'                        If llAmt = 0 Then
'                            tmProjByDay(ilStats).iNCUnits(ilWhichHour) = tmProjByDay(ilStats).iNCUnits(ilWhichHour) + 1     'total zero rate units
'                        End If
'
'                        If tmProjByDay(ilStats).iNCUnits(ilWhichHour) = 0 Then           'no NC exists yet, see if this spot is the lowest rate
'                            If tmProjByDay(ilStats).lMinRate(ilWhichHour) = 0 Then      'first time, set the lowest rate
'                                tmProjByDay(ilStats).lMinRate(ilWhichHour) = llAmt
'                            Else                                                        'lowest rate has been set so far, see if this spot is lower
'                                If llAmt < tmProjByDay(ilStats).lMinRate(ilWhichHour) Then        'get the lowest price in this hour
'                                    tmProjByDay(ilStats).lMinRate(ilWhichHour) = llAmt
'                                End If
'                            End If
'                        End If
'                        If llAmt > tmProjByDay(ilStats).lMaxRate(ilWhichHour) Then      'get the Highest price in this hour
'                            tmProjByDay(ilStats).lMaxRate(ilWhichHour) = llAmt
'                        End If
'
'                        tmProjByDay(ilStats).lSchedRev(ilWhichHour) = tmProjByDay(ilStats).lSchedRev(ilWhichHour) + llAmt   'schedule revenue
'                        tmProjByDay(ilStats).iSchedUnits(ilWhichHour) = tmProjByDay(ilStats).iSchedUnits(ilWhichHour) + 1   'scheduled units
'
'                        If tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "M" Then    'missed or cancelled
'                            tmProjByDay(ilStats).lMissedRev(ilWhichHour) = tmProjByDay(ilStats).lMissedRev(ilWhichHour) + llAmt   'schedule revenue
'                            tmProjByDay(ilStats).iMissedUnits(ilWhichHour) = tmProjByDay(ilStats).iMissedUnits(ilWhichHour) + 1   'schedule revenue
'                        End If
'                    End If                  'ilspotOK
'                End If
'            Next ilStats
'        End If
'    Next ilSpots
'    Exit Sub
'End Sub
