VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Begin VB.Form ExportSalesForce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Sales Force"
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
   Icon            =   "ExportSalesForce.frx":0000
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
      Picture         =   "ExportSalesForce.frx":08CA
      ScaleHeight     =   6015
      ScaleWidth      =   7485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7545
      Begin VB.Timer tmcSetTime 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   1155
         Top             =   5370
      End
      Begin VB.CheckBox ckcSpot 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Spot Data"
         Height          =   210
         Left            =   3360
         TabIndex        =   24
         Top             =   1200
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox ckcAux 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Auxillary Data"
         Height          =   210
         Left            =   3360
         TabIndex        =   23
         Top             =   840
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ListBox lbcInfo 
         Height          =   900
         Index           =   1
         ItemData        =   "ExportSalesForce.frx":A217
         Left            =   570
         List            =   "ExportSalesForce.frx":A219
         TabIndex        =   22
         Top             =   2190
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.ListBox lbcInfo 
         Height          =   1950
         Index           =   0
         ItemData        =   "ExportSalesForce.frx":A21B
         Left            =   570
         List            =   "ExportSalesForce.frx":A21D
         TabIndex        =   21
         Top             =   2445
         Width           =   6255
      End
      Begin VB.CommandButton cmcTo 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox plcTo 
         Height          =   375
         Index           =   1
         Left            =   1170
         ScaleHeight     =   315
         ScaleWidth      =   4245
         TabIndex        =   18
         Top             =   3375
         Visible         =   0   'False
         Width           =   4305
         Begin VB.TextBox edCTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   15
            Visible         =   0   'False
            Width           =   4230
         End
      End
      Begin VB.PictureBox plcTo 
         Height          =   375
         Index           =   0
         Left            =   1170
         ScaleHeight     =   315
         ScaleWidth      =   4245
         TabIndex        =   10
         Top             =   2955
         Visible         =   0   'False
         Width           =   4305
         Begin VB.TextBox edCTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   15
            TabIndex        =   12
            Top             =   15
            Visible         =   0   'False
            Width           =   4230
         End
      End
      Begin VB.CommandButton cmcTo 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
         Height          =   285
         Index           =   0
         Left            =   5640
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1200
      End
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   1500
         ItemData        =   "ExportSalesForce.frx":A21F
         Left            =   5625
         List            =   "ExportSalesForce.frx":A221
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   1140
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox ckcAll 
         BackColor       =   &H00FFFFC0&
         Caption         =   "All Vehicles"
         Height          =   210
         Left            =   5640
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox edcStart 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   3
         Top             =   810
         Width           =   1200
      End
      Begin VB.TextBox edcEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   300
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1200
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
         Height          =   495
         Index           =   0
         Left            =   570
         TabIndex        =   17
         Top             =   4000
         Visible         =   0   'False
         Width           =   6255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   1
         Left            =   570
         TabIndex        =   16
         Top             =   4005
         Visible         =   0   'False
         Width           =   6255
      End
      Begin DZACTXLibCtl.dzactxctrl zpcDZip 
         Left            =   705
         OleObjectBlob   =   "ExportSalesForce.frx":A223
         Top             =   4710
      End
      Begin VB.Label lacContract 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Contract #"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   240
         TabIndex        =   15
         Top             =   1680
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
         Left            =   240
         TabIndex        =   13
         Top             =   1200
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
         Left            =   3090
         TabIndex        =   11
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
         TabIndex        =   9
         Top             =   4860
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
Attribute VB_Name = "ExportSalesForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExportSalesForce.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text

Dim imRunningInBackGround As Integer            'tell ftp routine if running in background and to log in file vs on-screen

Dim smExportName(0 To 8) As String     '0 = spots, 1 = adv, 2 = agy, 3= prod prot, 4 = vehicle, 5 = slsp, 6 = sales source, 7 = sales office, 8 = contract
Dim hmFileHandle(0 To 8) As Integer       '0 = spots, 1 = adv, 2 = agy, 3= prod prot, 4 = vehicle, 5 = slsp, 6 = sales source, 7 = sales office, 8 = contract
Dim smExportSpotName As String
Dim imFirstActivate As Integer
Dim lmCntrNo As Long    'for debugging purposes to filter a single contract
Dim lmCntrCode As Long  'contract code so header doesnt have to be read to match all spots
Dim lmSlspRevShare() As Long      'accum gross $ for slsp splits
Dim lmTotalGross As Long          'contract total gross (incl NTR)

Dim lmTodayDate As Long
Dim smLatestYear As String
Dim smLatestMonth As String
Dim smLatestDay As String
Dim smYear As String
Dim smMonth As String
Dim smDay As String
Dim smNowDate As String
Dim smNowYear As String
Dim smNowMonth As String
Dim smNowDay As String

Dim imSetAll As Integer
Dim imAllClicked As Integer

Dim hmAdf As Integer
Dim tmAdf As ADF                  'Advt  record image
Dim imAdfRecLen As Integer        'ADF record length

Dim hmAgf As Integer
Dim tmAgf As AGF                  'AGency  record image
Dim imAgfRecLen As Integer        'AGF record length

Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim tmChfSrchKey As LONGKEY0

Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim tmClfSrchKey As CLFKEY0

Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF

Dim hmSbf As Integer            'Special billing (NTR)  file handle
Dim imSbfRecLen As Integer      'SBF record length
Dim tmSbf As SBF
Dim tmSbfSrchKey0 As SBFKEY0

Dim hmCef As Integer            'Comments  file handle
Dim imCefRecLen As Integer      'Comments record length
Dim tmCef As CEF
Dim tmCefSrchKey As LONGKEY0

Dim hmMnf As Integer
Dim tmMnf As MNF                  'Multiname  record image
Dim tmMnfSrchKey As INTKEY0
Dim tmMnfVGroup() As MNF
Dim tmMnfCompCode() As MNF
Dim tmMnfItem() As MNF
Dim imMnfRecLen As Integer        'MNF record length

Dim hmSdf As Integer
Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey3 As LONGKEY0
Dim tmSdfSrchKey4 As SDFKEY4            'date, chfcode

Dim hmSmf As Integer            'Spot MG file handle
Dim tmSmf As SMF                'Spot MG record image
Dim imSmfRecLen As Integer        'Spt MG record length
Dim tmSmfSrchKey2 As LONGKEY0   'SdfCode

Dim hmSlf As Integer
Dim tmSlf As SLF                  'Salesperson  record image
Dim imSlfRecLen As Integer        'Slsp record length

Dim hmSof As Integer
Dim tmSof As SOF                  'Slsp office  record image
Dim imSofRecLen As Integer        'SOF record length
Dim tmSofArray() As SOF

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length

Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer       'VSF record length

Dim tmSaf As SAF
Dim hmSaf As Integer
Dim imSafRecLen As Integer

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim lmNowDate As Long
Dim lmUserStartDate As Long         'user entered start date
Dim lmUserEndDate As Long           'user entered end date
Dim lmLatestSDFDate As Long         'Latest sdf spot date (for tfn)
Dim imAutoRun As Integer            '1 = auto run flag, 0 = manual
Dim imClosed As Integer

Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT
Dim tmSF_Output_Info As SF_OUTPUT_INFO

Private Type SPOTHEADER
    sContractID As String * 10          'chfcode
    sContract As String * 9             'chfcntrno
    sRev As String * 5                  'chfcntrevno
    sStartDate As String * 10           'chfstartdate
    sEndDate As String * 10             'chfenddate
    sGross As String * 11               'total gross
    sPctTrade As String * 3             '% of trade 0-100
    sOrderType As String * 1            'chftype
    sAgfCode As String * 5              'agency code
    sAdfCode As String * 5              'advt code
    sProduct As String * 35
    sPrimComp As String * 5             'prim comp code
    sSecComp As String * 5              'second comp code
    sSSCode As String * 5               'sales source code
    sSofCode As String * 5              'selling office code
    sSlfCode1 As String * 5             'splitslsp code #1
    sSlfCode2 As String * 5             'splitslsp code #2
    sSlfCode3 As String * 5             'splitslsp code #3
    sSlfGross1 As String * 11
    sSlfGross2 As String * 11
    sSlfGross3 As String * 11
End Type

Private Type VGINFO
    iCode As Integer
    sVehicleName As String
    sVehicleCode As String
    sVG2SubT As String
    sVG3Mkt As String
    sVG4Fmt As String
    sVG5Rsch As String
    sVG6SubC As String
End Type


Private Type SPOTDETAIL
    sSpotID As String * 10
    sSpotDate As String * 8
    sSpotTimeAired As String * 10
    sSpotLength As String * 3
    sVefCode As String * 5
    sVG2SubT As String * 20
    sVG3Mkt As String * 20
    sVG4Fmt As String * 20
    sVG5Rsch As String * 20
    sVG6SubC As String * 20
    sContractID As String * 10
    sAirTimeNTR As String * 1           'A = airtime, N = NTR
    sNTRType As String * 20             'ntr description
    sGrossSpotPrice As String
    sAgyCommPct As String

End Type
    
Private Type SF_OUTPUT_INFO
    iWhichFile As Integer
    iTotalFields As Integer
    sStr(0 To 25) As String         'max 25 string fields to write in 1 record, index zero ignored
End Type
 
Dim tmSpotDetail As SPOTDETAIL
Dim tmSpotHeader As SPOTHEADER
Dim tmVGINFO() As VGINFO            'vehicle group descriptions

Private Const DELIMITER = ","          'VERIFY WHAT  SHOULD BE: | vertical bar, or pipe character
 '0 = spots, 1 = adv, 2 = agy, 3= prod prot, 4 = vehicle, 5 = slsp, 6 = sales source, 7 = sales office, 8 = contract header
Private Const FILE_SPOTS = 0
Private Const FILE_ADV = 1
Private Const FILE_AGY = 2
Private Const FILE_PRODPROT = 3
Private Const FILE_VEHICLE = 4
Private Const FILE_SLSP = 5
Private Const FILE_SS = 6
Private Const FILE_OFFICE = 7
Private Const FILE_CONTRACT = 8

Private Sub edcStart_Click()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edCTo_Click(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcTo_GotFocus(ilIndex As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
'    gShowBranner
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        If imAutoRun = 0 Then
            MsgBox "Only one copy of ExportSalesForce can be run at a time ", vbInformation + vbOKOnly, "Counterpoint"
        End If
        gLogMsg "Second copy of ExportSalesForce path: " & App.Path & " from " & Trim$(gGetComputerName()), "ExportSalesForce.txt", False
        End
    End If
    gSetBkgdMode
    'igBkgdProg = 6
    mInit
    If imAutoRun = 0 Then                       'manual
        Me.WindowState = vbNormal
        If (Asc(tgSaf(0).sFeatures1) And SALESFORCEEXPORT) <> SALESFORCEEXPORT Then
            'Sales force now allowed
            gLogMsg "** Sales Force Export Disabled:  " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            If imAutoRun = 0 Then           'manual mode, show disallowed on screen
                lbcInfo(0).AddItem "Sales Force Export Disabled"
            End If
            imTerminate = True
            Exit Sub
        Else
            lacMin.Enabled = True
        End If
    Else                                        '1 = auto run
        Me.WindowState = vbMinimized
        If (Asc(tgSaf(0).sFeatures1) And SALESFORCEEXPORT) <> SALESFORCEEXPORT Then
            'Sales force now allowed
            gLogMsg "** Sales Force Export Disabled:  " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            imTerminate = True
        Else
            gOpenTmf
            tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
            tmcSetTime.Enabled = True
        End If
    End If
    tmcClick.Interval = 2000    '2 seconds
    tmcClick.Enabled = True
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
    If imAutoRun = 1 Then           'ilok = no errors creating the files; 1 = auto run
        tmcSetTime.Enabled = False
        gCloseTmf
    End If
    
    mCloseSpotFiles

    btrStopAppl
    Set ExportSalesForce = Nothing   'Remove data segment
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
    
    '************ change to 0 to DEBUG
    imAutoRun = 1                  '0 = user input, 1 = auto run
    ilPos = InStr(1, slCommand, "/UserInput", 1)
    If ilPos > 0 Then
        imAutoRun = 0
        igBkgdProg = 0
       ' gLogMsg "Current Mode: UserInput.", "ExportSalesForce.txt", False
    Else
        igBkgdProg = 14
'        gLogMsg "Current Mode: Background.", "ExportSalesForce.txt", False
    End If
    
    ' allow to run standalone in test
    slTestSystem = CurDir$
    If InStr(1, slTestSystem, "Test", vbTextCompare) = 0 Then
        ilTestSystem = False
    Else
        ilTestSystem = True
    End If
    
    mTestPervasive      'this is required to intiate the pvsw engine; otherwise err 3012 occurs with spfread
    gInitStdAlone ExportSalesForce, slStr, ilTestSystem
    
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "ExportProj.Frm"
    
    If igBkgdProg = 0 Then
        gLogMsg "Current Mode: UserInput.", "ExportSalesForce.txt", False
    Else
        gLogMsg "Current Mode: Background.", "ExportSalesForce.txt", False
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
    If Not gRecLengthOk("Slf.Btr", Len(tmSlf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Sof.Btr", Len(tmSof)) Then
        mCheckRecLength = False
        Exit Function
    End If
     If Not gRecLengthOk("ADf.Btr", Len(tmAdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
     If Not gRecLengthOk("Agf.Btr", Len(tmAgf)) Then
        mCheckRecLength = False
        Exit Function
    End If
     If Not gRecLengthOk("Mnf.Btr", Len(tmMnf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function

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
'
'               Begin processing to extract Spots and supporting files which include:
'               Advt, Agy, Vehicles, Product Protection, Salesperson,
'
'           THIS IS MANUAL MODE
Private Sub lacMin_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTodayDate As String
    Dim slLatestDate As String
    Dim slLatestMonth As String
    Dim slLatestDay As String
    Dim slLatestYear As String
    Dim slDefaultFileName As String
    Dim slDefaultSpotFileName As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDateTime As String
    Dim blAllOK As Boolean
    Dim ilExportLoop As Integer
    Dim slStartDate As String
    Dim blInitOutputErr As Boolean
    
    
    If imAutoRun = 0 Then           'manual mode
        'check for validity of date in case it was changed
        slStartDate = edcStart.Text
        If Not gValidDate(slStartDate) Then
            edcStart.SetFocus
            Exit Sub
        End If
        lmUserStartDate = gDateValue(slStartDate)
        
        slLatestDate = edcEnd.Text
        If Not gValidDate(slLatestDate) Then
            edcEnd.SetFocus
            Exit Sub
        End If
        
        'test start date is prior to end date
        lmUserEndDate = gDateValue(slLatestDate)
                
        If lmUserStartDate > lmUserEndDate Then
            edcEnd.SetFocus
            Exit Sub
        End If
        
'        gObtainYearMonthDayStr slStartDate, True, smYear, smMonth, smDay
'        edcStart.Text = slStartDate
    
        slLatestDate = Format$(lmUserEndDate, "m/d/yy")
        gObtainYearMonthDayStr slLatestDate, True, smLatestYear, smLatestMonth, smLatestDay
        
        If (Not blInitOutputErr) Or (Not imExporting) Then
            Screen.MousePointer = vbHourglass
            blInitOutputErr = mInitOutputFiles
            If Not blInitOutputErr Then                  'if any errors creating the files, terminate
                blAllOK = mExportAll()
                Screen.MousePointer = vbDefault
                If blAllOK Then
                    ilRet = gFTPMain(imRunningInBackGround)
                    lacStop.Caption = "Done"
                    lacMin.Enabled = False
                End If
            Else
                imTerminate = True
                Exit Sub
            End If
        End If
        
    Else
        Me.WindowState = vbMinimized
    End If
    Exit Sub
lacMinErr:
    ilRet = 1
    Resume Next
End Sub
Private Sub lacStop_Click()
Dim ilExportLoop As Integer
Dim hlFile As Integer

'   CANCEL button
'
    If Not imExporting Then
        imTerminate = True
        mTerminate
    Else
        imTerminate = True
        
        Screen.MousePointer = vbDefault
        imExporting = False
        lbcInfo(0).AddItem "Export Cancelled "
        lacStop.Caption = "Done"
        gLogMsg "Export Cancelled", "ExportSalesForce.txt", False
        mTerminate
    End If
End Sub


Private Sub mBtrErrorMsg(slMsg As String, ilInRet As Integer)
    Dim ilRet As Integer
    
    ilRet = ilInRet
    If ilRet >= 30000 Then
        ilRet = csiHandleValue(0, 7)
    End If
    gLogMsg slMsg & Trim$(Str$(ilRet)), "ExportSalesForce.txt", False
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
    ilRet = gPopUserVehicleBox(ExportSalesForce, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH + VEHSPORT, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)

    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExportSalesForce
        On Error GoTo 0
    End If
    
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'
'
'
'           mCloseSpotFiles - Close all applicable files for
'                       projection Export
'
Private Sub mCloseSpotFiles()
    Dim ilRet As Integer
    On Error Resume Next
    
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSbf)

    btrDestroy hmVef
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmSdf
    btrDestroy hmVsf
    btrDestroy hmSmf
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmSof
    btrDestroy hmAdf
    btrDestroy hmAgf
    btrDestroy hmMnf
    btrDestroy hmSbf
    
    Erase hmFileHandle, smExportName
    Erase tmSdfExtSort, tmSdfExt
    Erase tmMnfVGroup, tmMnfItem
    Erase tmMnfCompCode, tmSofArray
    Erase lmSlspRevShare, tmVGINFO

    Exit Sub
End Sub

'
'
'           mOpenFiles - open files applicable to Export Inventory and Spots
'
'
Private Function mOpenSpotFiles() As Integer
    Dim ilRet As Integer
    Dim ilTemp As Integer
    Dim ilError As Integer
    Dim slStamp As String
    Dim slMnfStamp As String

    ilError = False

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen VEF)", ExportSalesForce
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen CHF)", ExportSalesForce
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen CLF)", ExportSalesForce
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen CFF)", ExportSalesForce
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen SMF)", ExportSalesForce
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen VSF)", ExportSalesForce
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen SLF)", ExportSalesForce
    On Error GoTo 0
    imSlfRecLen = Len(tmSlf)

    
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen SOF)", ExportSalesForce
    On Error GoTo 0
    imSofRecLen = Len(tmSof)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen ADF)", ExportSalesForce
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen AGF)", ExportSalesForce
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
     
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen SMF)", ExportSalesForce
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    
    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenSpotFiles (btrOpen SBF)", ExportSalesForce
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf)

    gObtainAdvt
    gObtainAgency
    gObtainVef
    ilRet = gObtainSalesperson()
    
    'MNF opened in mInit for the client abbreviated name
    ilRet = gObtainMnfForType("C", slMnfStamp, tmMnfCompCode())
    If ilRet = False Then
        ilError = True
        Exit Function
    End If

    ilRet = gObtainMnfForType("H", slMnfStamp, tmMnfVGroup())
    If ilRet = False Then
        ilError = True
        Exit Function
    End If
    
    ilRet = gObtainMnfForType("I", slMnfStamp, tmMnfItem())         'NTR item types
    If ilRet = False Then
        ilError = True
        Exit Function
    End If


    mPopSof hmSof, tmSofArray()             'populate the Sales Offices
    
    mOpenSpotFiles = ilError
    Exit Function

mOpenSpotFilesErr:
    ilError = True
    
    Return
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
        llRet = SendMessageByNum(lbcVehicle.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        imAllClicked = False
    End If
    mSetCommands
End Sub
'Private Sub cmcTo_Click(ilIndex As Integer)
'    CMDialogBox.DialogTitle = "Export To File"
'    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
'    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
'    CMDialogBox.DefaultExt = ".Csv"
'    CMDialogBox.flags = cdlOFNCreatePrompt
'    CMDialogBox.Action = 1 'Open dialog
'    edCTo(0).Text = CMDialogBox.fileName
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'        ChDir sgCurDir
'    End If
'    If ilIndex = 0 Then
'        If edCTo(0).Text = "" Then
'            edCTo(0).Text = smExportAuxName
'        End If
'    Else
'        If edCTo(1).Text = "" Then
'            edCTo(1).Text = smExportSpotName
'        End If
'    End If
'End Sub
Private Sub edcContract_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcEnd_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcStart_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcTo_Change(ilIndex As Integer)
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
    
    ilRet = gObtainSAF()                    'get SAF to see if user OK to use feature
    
    If imAutoRun = 0 Then                   'manual
        lacMin.Enabled = False
        lacMin.Caption = "Export"
        imRunningInBackGround = False
    Else                                    'auto run
        lacMin.Caption = "Minimize"
        imRunningInBackGround = False
    End If
    imClosed = False
    imTerminate = False
    imFirstActivate = True
 
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
    ckcAll.Value = vbChecked            'vehicle list box is hidden, and all vehicle selected
    
    
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen MNF)", ExportSalesForce
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)

    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen SDF)", ExportSalesForce
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    
    mGetLatestSDFDate               'get the latest date in SDF to show as default
    edcEnd.Text = Format$(lmLatestSDFDate, "m/d/yy")
    
    Screen.MousePointer = vbDefault

    Exit Sub
mInitErr:
    Return
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
    'auxillary files have been closed after each process
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Traffic
    Unload ExportSalesForce
    'Set ExportSalesForce = Nothing   'Remve data segment
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
'
'                           tmcClick_Timer - AUTO RUN MODE
'
Private Sub tmcClick_Timer()
    Dim slRepeat As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slDateTimeSpot As String
    Dim slDefaultFileName As String
    Dim slDefaultSpotFileName As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slStartDate As String
    Dim slLatestSDFDate As String
    Dim slLatestMonth As String
    Dim slLatestDay As String
    Dim slLatestYear As String
    Dim blAllOK As Boolean
    Dim blInitOutputErr As Boolean
    
        tmcClick.Enabled = False
        If imTerminate Then
            lacStop_Click
            Exit Sub
        End If
        
        mCheckGG
        If (igGGFlag = 0) And (igRptGGFlag = 0) Then
            imTerminate = True
            lacStop_Click
            Exit Sub
        End If
    
        
        'slStartDate = Format$(lmNowDate - 3, "m/d/yy")       'get todays date to default
        slStartDate = Format$(lmNowDate, "m/d/yy")       'get todays date to default
        gObtainYearMonthDayStr slStartDate, True, slYear, slMonth, slDay
        edcStart.Text = slStartDate
      
        slLatestSDFDate = Format$(lmLatestSDFDate, "m/d/yy")     'obtain date of latest spot
        gObtainYearMonthDayStr slLatestSDFDate, True, slLatestYear, slLatestMonth, slLatestDay
        
        'use todays date and date of last spot scheduled as the default dates
        lmUserStartDate = gDateValue(slStartDate)
        lmUserEndDate = lmLatestSDFDate
                           
        ilRet = 0
        lacMin.Enabled = True       'only set Export button enabled once the start date has been filled in
        
        If imAutoRun = 1 Then           'ilok = no errors creating the files; 1 = auto run
            Screen.MousePointer = vbHourglass
            blInitOutputErr = mInitOutputFiles
            If Not blInitOutputErr Then
                gUpdateTaskMonitor 1, "SFE"
                blAllOK = mExportAll()
                gUpdateTaskMonitor 2, "SFE"
                If blAllOK Then
                    ilRet = gFTPMain(imRunningInBackGround)
                    ilRet = ilRet
                    lacStop_Click           'auto run, task finished, unload and return to caller
                Else
                    imTerminate = True
                    Exit Sub
                End If
                Screen.MousePointer = vbDefault
            Else
                imTerminate = True
                 lacStop_Click          'auto run, task finished, unload and return to caller
                Exit Sub
            End If

        End If
    Exit Sub
cmcExportDupNameErr:
    ilRet = 1
    Resume Next
End Sub

Public Sub mSetCommands()
    If lbcVehicle.SelCount > 0 And edcStart.Text <> "" And edcEnd.Text <> "" Then
        'cmcExport.Enabled = True
        lacMin.Enabled = True
    Else
        'cmcExport.Enabled = False
        lacMin.Enabled = False
    End If
    
End Sub
'
'                   mexportAll - Begin processing the inventory and spots for the user defined dates
'                   either defaulted to today minus 3 days to end of spot file, or user entered dates if run manually
'
'                   This export can be a standalone to be run automatically or manually.  If run automatically, the dates default
'                   to today minus 3 days to end of spot file.
'                   2 files are created:  Inv-mmddyy-mmddyy.txt that contains inventory for every day in the spot file for
'                                         conventional, selling and game vehicles.
'                                         Spot-mmddyy-mmddyy.txt that contains all spots for same period defined above.
'                   Both files are placed in the Export folder defined in traffic.ini
'                   ExportSalesForce.txt is a message file placed under data\messages folder
'
Private Function mExportAll() As Boolean
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
    Dim ilUpper As Integer
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    ReDim tmTradeCnts(0 To 0) As Long
    Dim slExportName As String
    Dim ilAnyInvErrors As Integer
    Dim ilAnySpotErrors As Integer
    Dim ilAnyErrors As Integer
    Dim hmExportFile As Integer
    Dim ilOk As Integer
    Dim ilExportLoop As Integer
    Dim hlFile As Integer
    Dim blExportOK As Boolean
    Dim blOpenFilesOK As Boolean
    'lbcInfo(0).Visible = False

        If imExporting Then
            Exit Function
        End If
    
        
        lmCntrNo = 0                'ths is for debugging on a single contract
        slStr = ExportSalesForce!edcContract
        If slStr <> "" Then
            lmCntrNo = Val(slStr)
        End If
        
        gLogMsg "** Begin Sales Force export: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
    
        If imAutoRun = 0 Then           'manual, disable the export button
            lacMin.Enabled = False
        End If
        
        
        If mOpenSpotFiles() Then
            imTerminate = True
            blExportOK = False
            Exit Function
        End If
        
        blExportOK = mExportAdvertiser
        If Not blExportOK Then
            imTerminate = True
            gLogMsg "** Export Advertiser Update Error, Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            lbcInfo(0).AddItem "Export Advertiser Update Error - terminated"
            Exit Function
        End If

        blExportOK = mExportAgency
        If Not blExportOK Then
            imTerminate = True
            gLogMsg "** Export Agency Update Error, Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            lbcInfo(0).AddItem "Export Agency Update Error - terminated"
            Exit Function
        End If

        
        blExportOK = mExportProdProt
        If Not blExportOK Then
            imTerminate = True
            gLogMsg "** Export Product Protection Update Error, Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            lbcInfo(0).AddItem "Export Product Protection Update Error - terminated"
            Exit Function
        End If
        
        blExportOK = mExportVehicle
        If Not blExportOK Then
            imTerminate = True
            gLogMsg "** Export Vehicle Write Update, Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            lbcInfo(0).AddItem "Export Vehicle Update Error - terminated"
            Exit Function
        End If
        
        blExportOK = mExportSlsp
        If Not blExportOK Then
            imTerminate = True
            gLogMsg "** Export Salesperson Update Error, Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            lbcInfo(0).AddItem "Export Salespeople Update Error - terminated"
            Exit Function
        End If
        
        blExportOK = mExportOffice
        If Not blExportOK Then
            imTerminate = True
            gLogMsg "** Export Sales Office Update Error, Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            lbcInfo(0).AddItem "Export Sales Office UpdateError - terminated"
            Exit Function
        End If
        
        blExportOK = mExportSalesSource
        If Not blExportOK Then
            imTerminate = True
            gLogMsg "** Export Sales Source Update Error, Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
            lbcInfo(0).AddItem "Export Sales Source Update Error - terminated"
            Exit Function
        End If


        
        gLogMsg "** Exporting Contracts & Spots: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem "Exporting Contracts & Spots"
        Screen.MousePointer = vbHourglass
        imExporting = True
        
        'build array of the vehicles in the same order as tgMVef array, and get each vehicle group names
        mBuildVGForVehicles
        
        ilAnyErrors = False         'general error flag
        
        ilErrorStatus = 0
        If lmCntrNo <> 0 Then           'single contract for debugging
            tmChfSrchKey1.lCntrNo = lmCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_END_OF_FILE Or tmChf.lCntrNo <> lmCntrNo Then
                'exit, no contract found
                gLogMsg "Contract #" & Trim$(ExportSalesForce!edcContract) & " does not exist ", "ExportSalesForce.txt", False
                lbcInfo(0).AddItem "Contract #" & Trim$(ExportSalesForce!edcContract) & " does not exist"
               
                mCloseSpotFiles
                
                lbcInfo(0).Visible = True
                lacStop.Caption = "&Done"
                Screen.MousePointer = vbDefault
                imExporting = False
                
                If imAutoRun = 1 Then     '9-7-05 if auto run, task finished, unload and return to caller
                    'cmcCancel_Click
                    lacStop_Click
                End If
                Exit Function
            Else
                lmCntrCode = tmChf.lCode
            End If
        End If
        blExportOK = mExportSpots()
            
            If Not blExportOK Then     'did all vehicles get processed for inventory without any errors
                lbcInfo(0).AddItem "Spots Partially Completed, saved in " & smExportName(FILE_SPOTS)
                gLogMsg "** Export Spots Partially completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "ExportSalesForce.txt", False
                gLogMsg "** Export Spots saved in " & smExportName(FILE_SPOTS), "ExportSalesForce.txt", False
            Else
                lbcInfo(0).AddItem "All Sales Force Exports Completed"
                gLogMsg "** Export Sales Force Completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "ExportSalesForce.txt", False
            End If
            
            For ilExportLoop = 0 To UBound(hmFileHandle)
                hlFile = hmFileHandle(ilExportLoop)
                Close #hlFile
            Next ilExportLoop
            
    
            mCloseSpotFiles
            Screen.MousePointer = vbDefault
    
'        Else
'            lbcInfo(0).AddItem "Open File Error: Export Failed"
'            gLogMsg "** Export File Open error " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "ExportSalesForce.txt", False

        lbcInfo(0).Visible = True
        Screen.MousePointer = vbDefault
        imExporting = False
        
                
        'If imAutoRun = 1 Then     'auto run, task finished, unload and return to caller , do later after testing ftp option
       '     lacStop_Click
        'End If
        mExportAll = blExportOK
        If (ilAnyErrors) Or (ilAnyInvErrors) Or (ilAnySpotErrors) Then          'any errors encountered
            mExportAll = False
        End If
        'mCloseSpotFiles
            
        Exit Function
mExportErr:
        ilRet = Err.Number
        Resume Next

                              
End Function
'
'           mGetLatestSdfDate - Obtain the latest date in the SDF file for tfn processing
'           Use this date to default End date processing.
Private Sub mGetLatestSDFDate()

Dim slLatestDate As String
Dim ilRet As Integer
            gPackDate "12/31/2069", tmSdfSrchKey4.iDate(0), tmSdfSrchKey4.iDate(1)
            tmSdfSrchKey4.lChfCode = 0
            ilRet = btrGetLessOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slLatestDate
            lmLatestSDFDate = gDateValue(slLatestDate)          'show this date as default to end date
            Exit Sub
End Sub
'
'               mExportSpots - Obtain all contracts active from user selected start/end dates, or the default dates (today minus 3 days to end of spot file).
'               Go back 1 month from start date to catch makegoods for contracts expired.
'               Process one contract at a time & gather the spots by contract.
'               Create an export entry for every spot.
'
Private Function mExportSpots() As Boolean
Dim slCntrStatus As String
Dim slCntrType As String
Dim ilHOState As Integer
Dim slStartDate As String
Dim slEndDate As String
Dim tlChfAdvtExt() As CHFADVTEXT
Dim ilRet As Integer
Dim llContrCode As Long
Dim llLoopOnSpots As Long
Dim llSpotIndex As Long
Dim ilClf As Integer
Dim ilCurrentRecd As Integer
Dim ilInx As Integer
Dim slDate As String
Dim llAmount As Long
Dim ilAgfInx As Integer
Dim ilCommPct As Integer
Dim slStr As String
Dim llDate As Long
Dim ilFound As Integer
Dim llTotalGross As Long
Dim blError As Boolean
Dim blExportOK As Boolean
Dim hlContract As Integer
Dim hlSpot As Integer
Dim slStrAmt As String
Dim ilLoop As Integer

        If lmCntrCode > 0 Then
            gLogMsg "** Exporting Contract #" & Trim$(Str$(edcContract.Text)), "ExportSalesForce.txt", False
            'ReDim tlChfAdvtExt(1 To 2) As CHFADVTEXT
            ReDim tlChfAdvtExt(0 To 1) As CHFADVTEXT
            tlChfAdvtExt(0).lCode = lmCntrCode
        Else
            slStartDate = Format$(lmUserStartDate - 30, "m/d/yy")        'go back 1 month to catch any contracts expired with makegoods scheduled into following month
            slEndDate = Format$(lmUserEndDate, "m/d/yy")
            ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
            slCntrStatus = "HO"                 'statuses: hold, order, not unsch hold, uns order
            slCntrType = "CVTRQ"         'all types: PI, DR, etc.  except PSA(p) and Promo(m)
            ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)
            ilRet = gObtainCntrForDate(ExportSalesForce, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
        End If
        
        'obtain todays date:  prior to todays date needs to get the scheduled times from ssf avails
        slDate = Format$(gNow(), "m/d/yy")
        lmTodayDate = gDateValue(gAdjYear(slDate))  'get todays date
        
        slStartDate = Format$(lmUserStartDate, "m/d/yy")        'these are the actual spot gathering dates
        slEndDate = Format$(lmUserEndDate, "m/d/yy")
       
        hlContract = hmFileHandle(FILE_CONTRACT)
        hlSpot = hmFileHandle(FILE_SPOTS)
        
        blExportOK = True
        For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1        'loop on all contracts from the selected user entered dates, retrieve spots one contract at a time
            If tlChfAdvtExt(ilCurrentRecd).sType <> "S" And tlChfAdvtExt(ilCurrentRecd).sType <> "M" Then       'ignore psa(S) and promo(M)
              'read the entire contract in memory to avoid rereading for each spot
              mInitSpotFields
              llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
              ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())       'read in entire contract
    
              If Not ilRet Then
                  gLogMsg "** Unable to obtain Contract Code " & Trim$(Str$(llContrCode)), "ExportSalesForce.txt", False
                  lbcInfo(0).AddItem "Unable to obtain Contract Code " & Trim$(Str$(llContrCode))
                blExportOK = False

            Else
                'put together the header information.  Gather once per contract
                ReDim lmSlspRevShare(0 To 9) As Long                'init slsp split
                tmSpotHeader.sContractID = Trim$(Str(tgChf.lCode))
                tmSpotHeader.sContract = Trim$(Str$(tgChf.lCntrNo))
                tmSpotHeader.sRev = Trim$(Str(tgChf.iCntRevNo))
                gUnpackDate tgChf.iStartDate(0), tgChf.iStartDate(1), slDate
                tmSpotHeader.sStartDate = Trim$(slDate)
                gUnpackDate tgChf.iEndDate(0), tgChf.iEndDate(1), slDate
                tmSpotHeader.sEndDate = Trim$(slDate)
                
                tmSpotHeader.sGross = ""                        'compute or header?
                tmSpotHeader.sPctTrade = Trim$(Str(tgChf.iPctTrade))
                tmSpotHeader.sOrderType = tgChf.sType        'regular, reservation, remnant, DR, PI
                tmSpotHeader.sAdfCode = Trim$(Str(tgChf.iAdfCode))
                tmSpotHeader.sAgfCode = Trim$(Str(tgChf.iAgfCode))
                tmSpotHeader.sProduct = Trim$(tgChf.sProduct)
                tmSpotHeader.sPrimComp = Trim$(Str(tgChf.iMnfComp(0)))        'primary competitive code
                tmSpotHeader.sSecComp = Trim$(Str(tgChf.iMnfComp(0)))          'secondary competitive code
                mCalcTotalGrossForLines                    'retain line total gross, need to add NTR to it
                
                tmSpotHeader.sSlfCode1 = Trim$(Str(tgChf.iSlfCode(0)))
                tmSpotHeader.sSlfCode2 = Trim$(Str(tgChf.iSlfCode(1)))
                tmSpotHeader.sSlfCode3 = Trim$(Str(tgChf.iSlfCode(2)))
                
                ilInx = gBinarySearchSlf(tgChf.iSlfCode(0))
                If ilInx >= 0 Then
                    tmSpotHeader.sSofCode = Trim$(Str(tgMSlf(ilInx).iSofCode))
                Else
                    tmSpotHeader.sSofCode = ""
                End If

                For ilLoop = LBound(tmSofArray) To (UBound(tmSofArray))
                    If tmSofArray(ilLoop).iCode = tgMSlf(ilInx).iSofCode Then
                        tmSpotHeader.sSSCode = Trim$(Str(tmSofArray(ilLoop).iMnfSSCode))
                        Exit For
                    End If
                Next ilLoop
               
                'gather all the contracts spots for todays date thru either end of spot file or user requested end date
                'sort by line (vs vehicle)
                ilRet = gObtainCntrSpot(-1, False, llContrCode, -1, "S", slStartDate, slEndDate, tmSdfExtSort(), tmSdfExt(), 0, False)    'find all contract spots, return by line sort
                mProcessSpots hlSpot
                
                'need to get the NTR $ before writing record
                blError = mExportNTR()
                
                mStripCents lmTotalGross, slStrAmt              'contract total (including NTR)
                tmSpotHeader.sGross = Trim$(slStrAmt)
                mStripCents lmSlspRevShare(0), slStrAmt
                tmSpotHeader.sSlfGross1 = Trim$(slStrAmt)
                mStripCents lmSlspRevShare(1), slStrAmt
                tmSpotHeader.sSlfGross2 = Trim$(slStrAmt)
                mStripCents lmSlspRevShare(2), slStrAmt
                tmSpotHeader.sSlfGross3 = Trim$(slStrAmt)
                
                blError = mWriteExportSpots(hlContract, FILE_CONTRACT)
                If blError Then
                    blExportOK = False
                    Exit For
                End If
              End If
            End If
        Next ilCurrentRecd
        mInsertOutFileNameMsg blExportOK, FILE_SPOTS
        mInsertOutFileNameMsg blExportOK, FILE_CONTRACT

        On Error Resume Next
        Erase tlChfAdvtExt
        Erase tmSdfExtSort
        Erase tmSdfExt
        Close #hlContract, #hlSpot
        
        mExportSpots = blExportOK
        
        Exit Function
End Function

'                   mWriteExportSpots - create a record for every spot from requested dates
'                   Create file Spot-mmddyy-mmddyy.txt (pipe character delimited)
'

Private Function mWriteExportSpots(hlFile As Integer, ilFileID As Integer) As Boolean

    Dim ilLoop As Integer
    Dim slRecord As String
    Dim blError As Boolean
    Dim ilRecdLen As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim hlSpots As Integer
    
            blError = False
            
            If ilFileID = FILE_SPOTS Then
                slRecord = Trim$(tmSpotDetail.sSpotID)
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sSpotDate) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sSpotTimeAired) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sSpotLength) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sVefCode) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sVG3Mkt) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sVG5Rsch) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sVG6SubC) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sVG4Fmt) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sVG2SubT) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sContractID) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sAirTimeNTR) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sNTRType) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sGrossSpotPrice) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotDetail.sAgyCommPct) & """"
            Else            'FILE_CONTRACT
                slRecord = Trim$(tmSpotHeader.sContractID)
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sContract) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sRev) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sStartDate) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sEndDate) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sGross) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sPctTrade) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sOrderType) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sAgfCode) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sAdfCode) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sProduct) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sPrimComp) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSecComp) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSSCode) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSofCode) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSlfCode1) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSlfCode2) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSlfCode3) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSlfGross1) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSlfGross2) & """"
                slRecord = slRecord & DELIMITER & """" & Trim$(tmSpotHeader.sSlfGross3) & """"
            End If
            
            ilRecdLen = Len(Trim(slRecord))
            ilRet = 0
            On Error GoTo mWriteExportSpotErr
            Print #hlFile, Left(slRecord, ilRecdLen)
            On Error GoTo 0
            If ilRet <> 0 Then
                blError = True
                imExporting = False
                gLogMsg "Error writing to Export Spot File, Error # " & Str$(ilRet), "ExportSalesForce.txt", False
                lbcInfo(0).AddItem "Erro writing to Export File: " & smExportName(ilFileID)

                mWriteExportSpots = False
                Exit Function
            End If
    
            mWriteExportSpots = blError
        Exit Function

mWriteExportSpotErr:
    On Error GoTo 0
    sgTmfStatus = "E"
    ilRet = Err.Number
    blError = True
    Exit Function
End Function

Private Sub mInitSpotFields()
 
        'initalize header information for a spot
        tmSpotHeader.sContractID = ""          'chfcode
        tmSpotHeader.sContract = ""            'chfcntrno
        tmSpotHeader.sRev = ""                  'chfcntrevno
        tmSpotHeader.sStartDate = ""          'chfstartdate
        tmSpotHeader.sEndDate = ""             'chfenddate
        tmSpotHeader.sGross = ""               'total gross
        tmSpotHeader.sPctTrade = ""             '% of trade 0-100
        tmSpotHeader.sOrderType = ""           'chftype
        tmSpotHeader.sAgfCode = ""            'agency code
        tmSpotHeader.sAdfCode = ""            'advt code
        tmSpotHeader.sProduct = ""
        tmSpotHeader.sPrimComp = ""            'prim comp code
        tmSpotHeader.sSecComp = ""              'second comp code
        tmSpotHeader.sSSCode = ""               'sales source code
        tmSpotHeader.sSofCode = ""             'selling office code
        tmSpotHeader.sSlfCode1 = ""             'splitslsp code #1
        tmSpotHeader.sSlfCode2 = ""             'splitslsp code #2
        tmSpotHeader.sSlfCode3 = ""            'splitslsp code #3
        tmSpotHeader.sSlfGross1 = ""
        tmSpotHeader.sSlfGross2 = ""
        tmSpotHeader.sSlfGross3 = ""
        
        
        'initialize detail spot information
        tmSpotDetail.sSpotID = ""
        tmSpotDetail.sSpotDate = ""
        tmSpotDetail.sSpotTimeAired = ""
        tmSpotDetail.sSpotLength = ""
        tmSpotDetail.sVefCode = ""
        tmSpotDetail.sContractID = ""
        tmSpotDetail.sAirTimeNTR = ""         'A = airtime, N = NTR
        tmSpotDetail.sNTRType = ""             'ntr description
        tmSpotDetail.sGrossSpotPrice = ""
        tmSpotDetail.sAgyCommPct = ""
        Exit Sub
End Sub

'
'                   mSearchCompCode - retrieve the competitive description from code obtained from contract header.
'                   Search thru table of competitive codes stored in memory
'                   <input> competitive code from header (primary or secondary)
'                   <return> competitive description
Private Function mSearchCompCode(ilCompCode As Integer) As String
Dim ilLoop As Integer

    mSearchCompCode = ""
    If ilCompCode = 0 Then
        Exit Function
    End If
    For ilLoop = LBound(tmMnfCompCode) To UBound(tmMnfCompCode) - 1
        If ilCompCode = tmMnfCompCode(ilLoop).iCode Then
            mSearchCompCode = tmMnfCompCode(ilLoop).sName
            Exit For
        End If
    Next ilLoop
    Exit Function
End Function

'
'                mpopsof - build into common array a table of all the Sales offices to
'               avoid rereading  for every contract
'               <input> SOF handle , file must be open
'               <return>  Array of all salesperson office records in
Private Sub mPopSof(hlSof As Integer, tlSof() As SOF)
Dim ilRet As Integer
Dim ilRecLen As Integer

    ReDim tlSof(0 To 0) As SOF
    ilRecLen = Len(tlSof(0))
    ilRet = btrGetFirst(hlSof, tlSof(UBound(tlSof)), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlSof(0 To UBound(tlSof) + 1) As SOF
        ilRet = btrGetNext(hlSof, tlSof(UBound(tlSof)), ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Sub
End Sub
'
'               mGetVGName - obtain all 5 vehicle groups defined for the vehicle, once per contract per vehicle.
'               Spots retrieved from contract are sorted by vehicle
'               vggroup1 = subtotals
'               vggroup2 = markets
'               vggroup3 = formats
'               vggroup4 = research
'               vggroup5 = subcompany
'               <input> vehicle group code from vehicle
'               Return - vehicle group description
Private Function mGetVGName(ilVGCode As Integer) As String
Dim ilLoop As Integer

        mGetVGName = ""
        If ilVGCode = 0 Then
            Exit Function
        End If
        For ilLoop = LBound(tmMnfVGroup) To UBound(tmMnfVGroup) - 1
            If tmMnfVGroup(ilLoop).iCode = ilVGCode Then
                mGetVGName = Trim$(tmMnfVGroup(ilLoop).sName)
                Exit For
            End If
        Next ilLoop

        Exit Function
End Function
'
'                   mGatherSpotDetail - get the information from spot to write the Detail spot records
'                   Create a record for each spot as its read
'                   <input>  hlSpot - spot table handle ID
'                   <output> none
Public Sub mProcessSpots(hlSpot As Integer)
Dim llAmount As Long
Dim llRecPos As Long
Dim ilRet As Integer
Dim ilCommPct As Integer
Dim slAmount As String
Dim slSharePct As String
Dim llNet As Long
Dim slStr As String
Dim llComm As Long
Dim ilDay As Integer
Dim slDate As String
Dim slTime As String
Dim slDayOfWeek As String
Dim tlCff As CFF
Dim llLoopOnSpots As Long
Dim llSpotIndex As Long
Dim llDate As Long
Dim ilClf As Integer
Dim ilInx As Integer
Dim ilSpotTime(0 To 1) As Integer
Dim ilAgyInx As Integer
Dim blExportErr As Boolean
Dim slStrAmt As String
Dim blFirstTime As Boolean

            blFirstTime = True
            For llLoopOnSpots = LBound(tmSdfExtSort) To UBound(tmSdfExtSort) - 1
                llSpotIndex = tmSdfExtSort(llLoopOnSpots).lSdfExtIndex
                gUnpackDateLong tmSdfExt(llSpotIndex).iDate(0), tmSdfExt(llSpotIndex).iDate(1), llDate
                If (llDate >= lmUserStartDate And llDate <= lmUserEndDate) And (tmSdfExt(llSpotIndex).sSchStatus = "S" Or tmSdfExt(llSpotIndex).sSchStatus = "O" Or tmSdfExt(llSpotIndex).sSchStatus = "G") And (tmSdfExt(llSpotIndex).sSpotType <> "X" And tmSdfExt(llSpotIndex).sSpotType <> "O" And tmSdfExt(llSpotIndex).sSpotType <> "C") Then          'valid spots must be within date span and normal sched type, mg or outside spot
                      
                    If (blFirstTime) Or (tmSdfExt(llSpotIndex).iLineNo <> tmClf.iLine) Then         'different line, find it
                          blFirstTime = False
                          For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                              tmClf = tgClf(ilClf).ClfRec
                              If tmSdfExt(llSpotIndex).iLineNo = tmClf.iLine Then
                                  Exit For
                              End If
                          Next ilClf
                    
                          'get vehicle information thats common to all spots in the same vehicle, once per vehicle
                          'but for outside and makegoods, they can be different vehicle than the vehicle of the line it came from,
                          'so retrieve vehicle info
                          ilInx = gBinarySearchVef(tmSdfExt(llSpotIndex).iVefCode)
                          If ilInx >= 0 Then
                              tmSpotDetail.sVG2SubT = Trim$(tmVGINFO(ilInx).sVG2SubT)
                              tmSpotDetail.sVG3Mkt = Trim$(tmVGINFO(ilInx).sVG3Mkt)
                              tmSpotDetail.sVG4Fmt = Trim$(tmVGINFO(ilInx).sVG4Fmt)
                              tmSpotDetail.sVG5Rsch = Trim$(tmVGINFO(ilInx).sVG5Rsch)
                              tmSpotDetail.sVG6SubC = Trim$(tmVGINFO(ilInx).sVG6SubC)
            
                          Else
                              tmSpotDetail.sVG2SubT = ""
                              tmSpotDetail.sVG3Mkt = ""
                              tmSpotDetail.sVG4Fmt = ""
                              tmSpotDetail.sVG5Rsch = ""
                              tmSpotDetail.sVG6SubC = ""
                              Exit For
                        End If
                    End If                  'tmsdfext(llspotindex).ilineno = tmclf.iline
                               
                    tmSpotDetail.sSpotID = Trim$(Str(tmSdfExt(llSpotIndex).lCode))                'internal spot id
                    gUnpackDate tmSdfExt(llSpotIndex).iDate(0), tmSdfExt(llSpotIndex).iDate(1), slDate
                    tmSpotDetail.sSpotDate = slDate  'spot date
    '                ilDay = gWeekDayStr(slDate)
    '                tmSpotDetail.sSpotDay = Mid$(slDayOfWeek, (ilDay * 3) + 1, 3)      'spot date
    '
                    'original scheduled time
                    ilSpotTime(0) = tmSdfExt(llSpotIndex).iTime(0)
                    ilSpotTime(1) = tmSdfExt(llSpotIndex).iTime(1)
    
                    gUnpackTime ilSpotTime(0), ilSpotTime(1), "A", "1", slTime         'time of avail in military time
                    tmSpotDetail.sSpotTimeAired = Trim$(slTime)
                    tmSpotDetail.sSpotLength = Trim$(Str$(tmSdfExt(llSpotIndex).iLen))    'spot length
                    tmSpotDetail.sVefCode = Trim$(Str$(tmSdfExt(llSpotIndex).iVefCode))     'internal vehicle code
                    tmSpotDetail.sAirTimeNTR = "A"
                    tmSpotDetail.sNTRType = ""              'ntr type
                    tmSpotDetail.sContractID = Trim$(Str$(tgChf.lCode))
                    llAmount = 0
                    llAmount = mGetSpotFlightNoSDF(tmSdfExt(llSpotIndex), hmSmf, tmSmf, tmClf, hmCff, tlCff)
                    
                    mStripCents llAmount, slStrAmt
                    tmSpotDetail.sGrossSpotPrice = Trim$(slStrAmt)
                    
                    If tgChf.iAgfCode = 0 Then          'direct
                        ilCommPct = 0                'no commission
                        slStrAmt = ".00"
                    Else
                        ilAgyInx = gBinarySearchAgf(tgChf.iAgfCode)
                        If ilAgyInx >= 0 Then
                            ilCommPct = tgCommAgf(ilAgyInx).iCommPct
                            slStrAmt = Trim$(gLongToStrDec(CLng(ilCommPct), 2))
                        End If
                    End If
                    tmSpotDetail.sAgyCommPct = Trim$(slStrAmt)
                    
                    blExportErr = mWriteExportSpots(hlSpot, FILE_SPOTS)
                    
                End If
            Next llLoopOnSpots

            Exit Sub
End Sub
'
'                   mInitOutputFiles - determine location to store output file.  Create each file and get its file handle for procesing
'                   Return = true if all filenames created
'
Private Function mInitOutputFiles() As Boolean
Dim ilExportLoop As Integer
Dim slDateTime As String
Dim ilRet As Integer
Dim slDate As String
Dim smClientName As String
Dim blInitOutputErr As Boolean
Dim slStr As String
Dim hlFile As Integer
Dim slIncrLetter As String * 1

        blInitOutputErr = False
        slDate = Format$(lmUserStartDate, "m/d/yy")
        gObtainYearMonthDayStr slDate, True, smYear, smMonth, smDay

        slDate = Format$(lmUserEndDate, "m/d/yy")
        gObtainYearMonthDayStr slDate, True, smLatestYear, smLatestMonth, smLatestDay

        slDate = Format$(gNow(), "m/d/yy")
        gObtainYearMonthDayStr slDate, True, smNowYear, smNowMonth, smNowDay

        smClientName = Trim$(tgSpf.sGClient)
        If tgSpf.iMnfClientAbbr > 0 Then
            tmMnfSrchKey.iCode = tgSpf.iMnfClientAbbr
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                smClientName = Trim$(tmMnf.sName)
            End If
        End If
        
        ReDim gFtpArray(0 To 8) As String           'files to ftp
        
        slIncrLetter = ""
        For ilExportLoop = LBound(hmFileHandle) To UBound(hmFileHandle)   '0 = spots, 1 = adv, 2 = agy, 3= prod prot, 4 = vehicle, 5 = slsp, 6 = sales source, 7 = sales office, 8 = contract
            ilRet = 0
            slIncrLetter = ""

           'Filenames:  SFORCE mmddyy ccccccc.txt
           'SFORCE : hard coded
           'all files have type of data appended  (i.e. SFORCE- Vehicle, SFORCE- Spots)
           'mmddyy date generated
           'ccccc - client short form name
           Do
                If ilExportLoop = 0 Then
                    smExportName(ilExportLoop) = sgExportPath & "SForce-Spots " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"     'this is requested dates
                    'smExportName(ilExportLoop) = sgExportPath & "SForce-Spots " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                ElseIf ilExportLoop = 1 Then
                    'smExportName(ilExportLoop) = sgExportPath & "SForce-Advertiser " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                    smExportName(ilExportLoop) = sgExportPath & "SForce-Advertiser " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                ElseIf ilExportLoop = 2 Then
                    'smExportName(ilExportLoop) = sgExportPath & "SForce-Agency " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                     smExportName(ilExportLoop) = sgExportPath & "SForce-Agency " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                ElseIf ilExportLoop = 3 Then
                     'smExportName(ilExportLoop) = sgExportPath & "SForce-Product Protection " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                     smExportName(ilExportLoop) = sgExportPath & "SForce-Product Protection " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                ElseIf ilExportLoop = 4 Then
                    'smExportName(ilExportLoop) = sgExportPath & "SForce-Vehicle " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                    smExportName(ilExportLoop) = sgExportPath & "SForce-Vehicle " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                ElseIf ilExportLoop = 5 Then
                    'smExportName(ilExportLoop) = sgExportPath & "SForce-Salesperson " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                    smExportName(ilExportLoop) = sgExportPath & "SForce-Salesperson " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                ElseIf ilExportLoop = 6 Then
                    'smExportName(ilExportLoop) = sgExportPath & "SForce-Sales Source " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                    smExportName(ilExportLoop) = sgExportPath & "SForce-Sales Source " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                ElseIf ilExportLoop = 7 Then
                    'smExportName(ilExportLoop) = sgExportPath & "SForce-Office " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                    smExportName(ilExportLoop) = sgExportPath & "SForce-Office " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                ElseIf ilExportLoop = 8 Then
                    'smExportName(ilExportLoop) = sgExportPath & "SForce-Contract " & (smNowMonth) & Trim$(smNowDay) & Trim$(Mid$(smNowYear, 3)) & " " & Trim$(smClientName) & ".txt"
                    smExportName(ilExportLoop) = sgExportPath & "SForce-Contract " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                End If
            
                
                If (InStr(smExportName(ilExportLoop), ":") = 0) And (Left$(smExportName(ilExportLoop), 2) <> "\\") Then     'test for absence of colon and not using \\
                    smExportName(ilExportLoop) = sgExportPath & smExportName(ilExportLoop)
                End If
            
                'ilRet = 0
                'On Error GoTo mInitOutputFiles:
                 'slDateTime = FileDateTime(smExportName(ilExportLoop))
                 ilRet = gFileExist(smExportName(ilExportLoop))
                 If ilRet = 0 Then      'err 53 file not found; 0 = file found
                    If slIncrLetter = " " Then
                        slIncrLetter = "A"
                    Else
                        slIncrLetter = Chr(Asc(slIncrLetter) + 1)
                        'Kill smExportName(ilExportLoop)
                    End If
                 End If
              Loop While ilRet = 0
             
                'hmFileHandle(ilExportLoop) = FreeFile
                ilRet = 0
                'Open smExportName(ilExportLoop) For Output As hmFileHandle(ilExportLoop)
                ilRet = gFileOpen(smExportName(ilExportLoop), "Output", hmFileHandle(ilExportLoop))
                On Error GoTo 0
                If ilRet <> 0 Then
                    gLogMsg "Open File Error #" & Str$(Err.Number) & smExportName(ilExportLoop), "ExportSalesForce.txt", False
                    Close #hmFileHandle(ilExportLoop)
                    imExporting = False
                    Screen.MousePointer = vbDefault
                    If imAutoRun = 0 Then
                        lbcInfo(0).AddItem "Open File Error #" & Str$(Err.Number) & " for " & smExportName(ilExportLoop)
                    End If
                    mInitOutputFiles = False          'error
                    Exit Function
                End If
        
           
               'write out the column headings
                hlFile = hmFileHandle(ilExportLoop)
                Select Case ilExportLoop
                    Case FILE_PRODPROT
                        slStr = "Code" & DELIMITER & "Product Protection Name"
                        gFtpArray(ilExportLoop) = "SForce-Product Protection " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                    Case FILE_VEHICLE
                        slStr = "Code" & DELIMITER & "Vehicle Name" & DELIMITER & "State" & DELIMITER & "Type"
                        gFtpArray(ilExportLoop) = "SForce-Vehicle " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                    Case FILE_SLSP
                        slStr = "Code" & DELIMITER & "Slsp First Name" & DELIMITER & "Slsp Last Name" & DELIMITER & "Email Address" & DELIMITER & "State" & DELIMITER & "Office Code"
                        gFtpArray(ilExportLoop) = "SForce-Salesperson " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                    Case FILE_SS
                        slStr = "Code" & DELIMITER & "Sales Source Name"
                        gFtpArray(ilExportLoop) = "SForce-Sales Source " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                    Case FILE_OFFICE
                        slStr = "Code" & DELIMITER & "Office Name" & DELIMITER & "Sales Source Code"
                        gFtpArray(ilExportLoop) = "SForce-Office " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                    Case FILE_ADV
                        slStr = "Code" & DELIMITER & "Advertiser Name" & DELIMITER & "State"
                        gFtpArray(ilExportLoop) = "SForce-Advertiser " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                    Case FILE_AGY
                        slStr = "Code" & DELIMITER & "Agency Name" & DELIMITER & "State"
                        gFtpArray(ilExportLoop) = "SForce-Agency " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                    Case FILE_CONTRACT
                         slStr = "ID" & DELIMITER & "Contract #" & DELIMITER & "Rev #" & DELIMITER & "Start Date" & DELIMITER & "End Date" & DELIMITER & "Total Gross" & DELIMITER & "Pct Trade" & DELIMITER & "Order Type" & DELIMITER & "Agency Code"
                         slStr = slStr & DELIMITER & "Advertiser Code" & DELIMITER & "Product Name" & DELIMITER & "1st Comp Code" & DELIMITER & "2nd Comp Code" & DELIMITER & "Sales Source Code" & DELIMITER & "Sales Office Code"
                         slStr = slStr & DELIMITER & "Salesperson1 Code" & DELIMITER & "Salesperson2 Code" & DELIMITER & "Salesperson3 Code" & DELIMITER & "Salesperson1 Split" & DELIMITER & "Salesperson2 Split" & DELIMITER & "Salesperson3 Split"
                         gFtpArray(ilExportLoop) = "SForce-Contract " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                    Case FILE_SPOTS
                        slStr = "Spot ID" & DELIMITER & "Spot Date" & DELIMITER & "Spot Time" & DELIMITER & "Length" & DELIMITER & "Vehicle Code" & DELIMITER & "Market" & DELIMITER & "Research" & DELIMITER & "Sub-Company" & DELIMITER & "Format"
                        slStr = slStr & DELIMITER & "SubTotal" & DELIMITER & "Contract ID" & DELIMITER & "Airtime/NTR" & DELIMITER & "NTR Type" & DELIMITER & "Gross Spot Price" & DELIMITER & "Agency Comm Pct"
                        gFtpArray(ilExportLoop) = "SForce-Spots " & (smMonth) & Trim$(smDay) & Trim$(Mid$(smYear, 3)) & "-" & Trim$(smLatestMonth) & Trim$(smLatestDay) & Trim$(Mid$(smLatestYear, 3)) & Trim$(slIncrLetter) & " " & Trim$(smClientName) & ".txt"
                End Select
                
                On Error GoTo mWriteColumnHeadingErr
                Print #hlFile, slStr     'write header description
                On Error GoTo 0
        
                slStr = "As of " & Format$(gNow(), "mm/dd/yy") & " "
                slStr = slStr & Format$(gNow(), "h:mm:ssAM/PM")
    
                On Error GoTo mWriteColumnHeadingErr
                Print #hlFile, slStr        'write header description
                On Error GoTo 0
   
       Next ilExportLoop
       
       mInitOutputFiles = blInitOutputErr
       Exit Function
       
'mInitOutputFiles:
'    ilRet = Err.Number
'    Resume Next
    
mWriteColumnHeadingErr:
    sgTmfStatus = "E"
    ilRet = Err.Number
    blInitOutputErr = True
    mInitOutputFiles = blInitOutputErr        'error
    gLogMsg "** Cannot write columnn heading in mInitOutputFiles, export Terminated: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
    lbcInfo(0).AddItem "Cannot write columnn heading in mInitOutputFiles, export Terminated"

    Exit Function
    
End Function
'
'                   mExportProdProt - read all Product Protection codes (competitives) and create the code and name
'                   Return : true for no errors
'
Private Function mExportProdProt() As Boolean

    Dim blExportOK As Boolean
    Dim ilRet As Integer
    Dim llLoopOnItem As Long
    Dim blError As Boolean
    Dim hlFile As Integer
    Dim tlMnf() As MNF
    
        On Error GoTo MExportProdProtErr:
        blExportOK = True
        gLogMsg "** Exporting Product Protection: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem "Exporting Product Protection"
        hlFile = hmFileHandle(FILE_PRODPROT)
        ilRet = gObtainMnfForType("C", sgCompMnfStamp, tlMnf())
        For llLoopOnItem = LBound(tlMnf) To UBound(tlMnf) - 1
            tmSF_Output_Info.iWhichFile = FILE_PRODPROT
            tmSF_Output_Info.sStr(1) = Trim$(Str(tlMnf(llLoopOnItem).iCode))
            tmSF_Output_Info.sStr(2) = Trim$(tlMnf(llLoopOnItem).sName)
            tmSF_Output_Info.iTotalFields = 2
            blError = mWriteExportAuxRec(hlFile, tmSF_Output_Info)
            If blError Then
                blExportOK = False
                Exit For
            End If
        Next llLoopOnItem
        mInsertOutFileNameMsg blExportOK, FILE_PRODPROT
        Close #hlFile
        mExportProdProt = blExportOK
        Exit Function
        
MExportProdProtErr:
    blExportOK = False
    
End Function
'
'               mWriteExportAuxRec - write the export file given the type to be created
'               <input> hlFile - file open handle
'                       tlSF_Output_Info() - array of 1 records fields to write
'
'
Private Function mWriteExportAuxRec(hlFile As Integer, tlSF_Output_Info As SF_OUTPUT_INFO) As Boolean
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilIndex As Integer
    Dim blError As Boolean
    Dim ilTotalFields As Integer
    Dim ilRet As Integer
    Dim ilWhichFile As Integer
    
        blError = False

        ilTotalFields = tlSF_Output_Info.iTotalFields
        For ilLoop = 1 To ilTotalFields
            If ilLoop = 1 Then
                slStr = """" & Trim$(tlSF_Output_Info.sStr(1)) & """"
            Else
                slStr = slStr & DELIMITER & """" & Trim$(tlSF_Output_Info.sStr(ilLoop)) & """"
            End If
        Next ilLoop
        On Error GoTo mWriteExportAuxRecErr
        Print #hlFile, slStr
        On Error GoTo 0

        mWriteExportAuxRec = blError
        Exit Function

mWriteExportAuxRecErr:
    On Error GoTo 0
    sgTmfStatus = "E"
    ilRet = Err.Number
    blError = True
    Exit Function

End Function
'
'                   mExportVehicle - read all vehicles nd create the code and name and status
'                   Return : true for no errors
'
Private Function mExportVehicle() As Boolean

    Dim blExportOK As Boolean
    Dim ilRet As Integer
    Dim llLoopOnItem As Long
    Dim blError As Boolean
    Dim hlFile As Integer
    
        blExportOK = True
        gLogMsg "** Exporting Vehicles: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem "Exporting Vehicles"
        hlFile = hmFileHandle(FILE_VEHICLE)
        For llLoopOnItem = LBound(tgMVef) To UBound(tgMVef) - 1
            tmSF_Output_Info.iTotalFields = 4
            tmSF_Output_Info.iWhichFile = FILE_VEHICLE
            tmSF_Output_Info.sStr(1) = Trim$(Str(tgMVef(llLoopOnItem).iCode))   'internal vehicle code
            tmSF_Output_Info.sStr(2) = Trim$(tgMVef(llLoopOnItem).sName)       'vehicle name
            tmSF_Output_Info.sStr(3) = Trim$(tgMVef(llLoopOnItem).sState)   'A = active, D = dormant
            tmSF_Output_Info.sStr(4) = Trim$(tgMVef(llLoopOnItem).sType)    'conventional, selling, airing, etc
            blError = mWriteExportAuxRec(hlFile, tmSF_Output_Info)
            If blError Then
                blExportOK = False
                Exit For
            End If
        Next llLoopOnItem
        mInsertOutFileNameMsg blExportOK, FILE_VEHICLE
        Close #hlFile
        mExportVehicle = blExportOK
        Exit Function
End Function
'
'                   mExportSlsp - read all salespeople and create the code and name , email addr, status and office codes
'                   Return : true for no errors
'
Private Function mExportSlsp() As Boolean

    Dim blExportOK As Boolean
    Dim ilRet As Integer
    Dim llLoopOnItem As Long
    Dim blError As Boolean
    Dim hlFile As Integer
    Dim ilLoopOnUser As Integer
    
        blExportOK = True
        gLogMsg "** Exporting Salespeople: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem "Exporting Salespeople"
        hmCef = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            blExportOK = False
            Exit Function
        End If

        hlFile = hmFileHandle(FILE_SLSP)
        For llLoopOnItem = LBound(tgMSlf) To UBound(tgMSlf) - 1
            tmSF_Output_Info.iWhichFile = FILE_SLSP
            tmSF_Output_Info.sStr(1) = Trim$(Str(tgMSlf(llLoopOnItem).iCode))   'internal slsp code
            tmSF_Output_Info.sStr(2) = Trim$(tgMSlf(llLoopOnItem).sFirstName)       'slsp first name
            tmSF_Output_Info.sStr(3) = Trim$(tgMSlf(llLoopOnItem).sLastName)       'slsp last name
            tmSF_Output_Info.sStr(4) = ""                       ' slsp email address
            For ilLoopOnUser = LBound(tgPopUrf) To UBound(tgPopUrf) - 1
                If tgMSlf(llLoopOnItem).iCode = tgPopUrf(ilLoopOnUser).iSlfCode Then
                    tmCefSrchKey.lCode = tgPopUrf(ilLoopOnUser).lEMailCefCode 'E-Mail reference pointer
                    ilRet = btrGetEqual(hmCef, tmCef, Len(tmCef), tmCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        tmSF_Output_Info.sStr(4) = Trim$(tmCef.sComment)
                    End If
                    Exit For
                End If
            Next ilLoopOnUser
            
            tmSF_Output_Info.sStr(5) = Trim$(tgMSlf(llLoopOnItem).sState)       'status, active/dormant
            tmSF_Output_Info.sStr(6) = Trim$(Str(tgMSlf(llLoopOnItem).iSofCode))       'office code
            tmSF_Output_Info.iTotalFields = 6
            blError = mWriteExportAuxRec(hlFile, tmSF_Output_Info)
            If blError Then
                blExportOK = False
                Exit For
            End If
        Next llLoopOnItem
        mInsertOutFileNameMsg blExportOK, FILE_SLSP
        Close #hlFile
        ilRet = btrClose(hmCef)
        btrDestroy hmCef
        mExportSlsp = blExportOK
        Exit Function
End Function
'
'                   mExportOffice - read all Sales Office codes  and create the code and name
'                   Return : true for no errors
'
Private Function mExportOffice() As Boolean

    Dim blExportOK As Boolean
    Dim ilRet As Integer
    Dim llLoopOnItem As Long
    Dim blError As Boolean
    Dim hlFile As Integer
    Dim tlMnf() As MNF
    
        blExportOK = True
        gLogMsg "** Exporting Sales Office: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem "Exporting Sales Offices"
        hlFile = hmFileHandle(FILE_OFFICE)
        For llLoopOnItem = LBound(tmSofArray) To UBound(tmSofArray) - 1
            tmSF_Output_Info.iWhichFile = FILE_OFFICE
            tmSF_Output_Info.sStr(1) = Trim$(Str(tmSofArray(llLoopOnItem).iCode))
            tmSF_Output_Info.sStr(2) = Trim$(tmSofArray(llLoopOnItem).sName)
            tmSF_Output_Info.sStr(3) = Trim$(Str(tmSofArray(llLoopOnItem).iMnfSSCode))       'sales source code
            tmSF_Output_Info.iTotalFields = 3
            blError = mWriteExportAuxRec(hlFile, tmSF_Output_Info)
            If blError Then
                blExportOK = False
                Exit For
            End If
        Next llLoopOnItem
        mInsertOutFileNameMsg blExportOK, FILE_OFFICE
        Close #hlFile
        mExportOffice = blExportOK
        Exit Function
End Function
'
'                   mExportSalesSource - read all Sales Source codes (competitives) and create the code and name
'                   Return : true for no errors
'
Private Function mExportSalesSource() As Boolean

    Dim blExportOK As Boolean
    Dim ilRet As Integer
    Dim llLoopOnItem As Long
    Dim blError As Boolean
    Dim hlFile As Integer
    Dim tlMnf() As MNF
    
        blExportOK = True
        gLogMsg "** Exporting Sales Source: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem "Exporting Sales Source"
        hlFile = hmFileHandle(FILE_SS)
        ilRet = gObtainMnfForType("S", sgCompMnfStamp, tlMnf())
        For llLoopOnItem = LBound(tlMnf) To UBound(tlMnf) - 1
            tmSF_Output_Info.iWhichFile = FILE_SS
            tmSF_Output_Info.sStr(1) = Trim$(Str(tlMnf(llLoopOnItem).iCode))
            tmSF_Output_Info.sStr(2) = Trim$(tlMnf(llLoopOnItem).sName)
            tmSF_Output_Info.iTotalFields = 2
            blError = mWriteExportAuxRec(hlFile, tmSF_Output_Info)
            If blError Then
                blExportOK = False
                Exit For
            End If
        Next llLoopOnItem
        mInsertOutFileNameMsg blExportOK, FILE_SS
        Close #hlFile
        mExportSalesSource = blExportOK
        Exit Function
End Function
'
'                   Build an array the same size as tgMVef (array of vehicles built into global array
'                   Get the vehicle group descriptions for each of the 5 groups
'                   mBuildVGForVehicles
'                   <input> none
'                   <output> Generated array of vehicle group descriptions
Private Sub mBuildVGForVehicles()
Dim ilLoop As Integer

        ReDim tmVGINFO(LBound(tgMVef) To UBound(tgMVef)) As VGINFO          'create same size array as the vehicle table created
        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1
            tmVGINFO(ilLoop).iCode = tgMVef(ilLoop).iCode
            tmVGINFO(ilLoop).sVehicleCode = Trim$(Str(tgMVef(ilLoop).iCode))
            tmVGINFO(ilLoop).sVehicleName = Trim$(tgMVef(ilLoop).sName)
            'if no vehicle group exists, return blank field
            tmVGINFO(ilLoop).sVG2SubT = mGetVGName(tgMVef(ilLoop).iMnfVehGp2)            'subtotals
            tmVGINFO(ilLoop).sVG3Mkt = mGetVGName(tgMVef(ilLoop).iMnfVehGp3Mkt)            'markets
            tmVGINFO(ilLoop).sVG4Fmt = mGetVGName(tgMVef(ilLoop).iMnfVehGp4Fmt)            'formats
            tmVGINFO(ilLoop).sVG5Rsch = mGetVGName(tgMVef(ilLoop).iMnfVehGp5Rsch)            'research
            tmVGINFO(ilLoop).sVG6SubC = mGetVGName(tgMVef(ilLoop).iMnfVehGp6Sub)            'subcompanies
        Next ilLoop
        Exit Sub
End Sub
'
'                   mExportAdvertiser - read all advertisers  and create the code and name
'                   Return : true for no errors
'
Private Function mExportAdvertiser() As Boolean

    Dim blExportOK As Boolean
    Dim ilRet As Integer
    Dim llLoopOnItem As Long
    Dim blError As Boolean
    Dim hlFile As Integer
    
        blExportOK = True
        gLogMsg "** : " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem "Exporting Advertisers"
        hlFile = hmFileHandle(FILE_ADV)
        For llLoopOnItem = LBound(tgCommAdf) To UBound(tgCommAdf) - 1
            tmSF_Output_Info.iWhichFile = FILE_ADV
            tmSF_Output_Info.sStr(1) = Trim$(Str(tgCommAdf(llLoopOnItem).iCode))
            tmSF_Output_Info.sStr(2) = Trim$(tgCommAdf(llLoopOnItem).sName)
            tmSF_Output_Info.sStr(3) = tgCommAdf(llLoopOnItem).sState      'active or dormant
            tmSF_Output_Info.iTotalFields = 3
            blError = mWriteExportAuxRec(hlFile, tmSF_Output_Info)
            If blError Then
                blExportOK = False
                Exit For
            End If
        Next llLoopOnItem
        mInsertOutFileNameMsg blExportOK, FILE_ADV
        Close #hlFile
        mExportAdvertiser = blExportOK
        Exit Function
End Function
'
'                   mExportAgency- read all agencies  and create the code and name
'                   Return : true for no errors
'
Private Function mExportAgency() As Boolean

    Dim blExportOK As Boolean
    Dim ilRet As Integer
    Dim llLoopOnItem As Long
    Dim blError As Boolean
    Dim hlFile As Integer
    
        blExportOK = True
        gLogMsg "** Exporting Agencies: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem "Exporting Agencies"
        hlFile = hmFileHandle(FILE_AGY)
        For llLoopOnItem = LBound(tgCommAgf) To UBound(tgCommAgf) - 1
            tmSF_Output_Info.iWhichFile = FILE_AGY
            tmSF_Output_Info.sStr(1) = Trim$(Str(tgCommAgf(llLoopOnItem).iCode))
            tmSF_Output_Info.sStr(2) = Trim$(tgCommAgf(llLoopOnItem).sName)
            tmSF_Output_Info.sStr(3) = tgCommAgf(llLoopOnItem).sState       'active or dormant
            tmSF_Output_Info.iTotalFields = 3
            blError = mWriteExportAuxRec(hlFile, tmSF_Output_Info)
            If blError Then
                blExportOK = False
                Exit For
            End If
        Next llLoopOnItem
        mInsertOutFileNameMsg blExportOK, FILE_AGY
        Close #hlFile
        mExportAgency = blExportOK
        Exit Function
End Function
'
'               mGetSpotFlightNoSDF - obtain the spot flight price without requiring SDF.
'               use entry from SDFEXT array.
Private Function mGetSpotFlightNoSDF(tlSdfExt As SDFEXT, hlSmf As Integer, tlSmf As SMF, tlClf As CLF, hlCff As Integer, tlCff As CFF) As Long
'
'   llAmount = mGetSpotFlight(tlSdfExt, tlClf, hlCff, hlSmf, tlCff)
'   Where:
'       tlSdfext(I)- Extended spot image from sdfext array
'       tlClf(I)- Line image to get iCntRevNo and iPropVer
'       hlCff(I)- Open Handle to CFF
'       tlCff(O)- Spot flight
'       llAmount = flight price, 0 if no flight found
'

    Dim llSpotDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim tlCffSrchKey As CFFKEY0 'CFF key record image
    Dim ilCffRecLen As Integer     'CFF record length
    Dim ilRet As Integer

    mGetSpotFlightNoSDF = 0
    If (tlSdfExt.sSchStatus = "G") Or (tlSdfExt.sSchStatus = "O") Then
       ilRet = mFindSmf(tlSdfExt, hlSmf, tlSmf)
       If ilRet = False Then
           Exit Function
       End If
       gUnpackDateLong tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), llSpotDate
    Else
        gUnpackDateLong tlSdfExt.iDate(0), tlSdfExt.iDate(1), llSpotDate
    End If
    tlCffSrchKey.lChfCode = tlSdfExt.lChfCode  'llChfCode
    tlCffSrchKey.iClfLine = tlClf.iLine 'tlSdf.iLineNo using line so avg price can be obtained for package line which bill by airing
    tlCffSrchKey.iCntRevNo = tlClf.iCntRevNo
    tlCffSrchKey.iPropVer = tlClf.iPropVer
    tlCffSrchKey.iStartDate(0) = 0
    tlCffSrchKey.iStartDate(1) = 0
    ilCffRecLen = Len(tlCff)
    ilRet = btrGetGreaterOrEqual(hlCff, tlCff, ilCffRecLen, tlCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlCff.lChfCode = tlSdfExt.lChfCode) And (tlCff.iClfLine = tlClf.iLine)
        If (tlCff.iCntRevNo = tlClf.iCntRevNo) And (tlCff.iPropVer = tlClf.iPropVer) Then 'And (tmCff(2).sDelete <> "Y") Then
            gUnpackDateLong tlCff.iStartDate(0), tlCff.iStartDate(1), llStartDate    'Week Start date
            gUnpackDateLong tlCff.iEndDate(0), tlCff.iEndDate(1), llEndDate    'Week Start date
            If (llSpotDate >= llStartDate) And (llSpotDate <= llEndDate) Then
                mGetSpotFlightNoSDF = tlCff.lActPrice
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hlCff, tlCff, ilCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gRemoveSmf                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove MG Spec for a spot      *
'*                                                     *
'*******************************************************
Private Function mFindSmf(tlSdfExt As SDFEXT, hlSmf As Integer, tlSmf As SMF) As Integer
'
'   ilRet = gRemoveSmf(tlSdf, hlSmf, tlSmf)
'   Where:
'       tlSdf(I/O)- Sdf record
'       hlSmf(I)- Smf open handle
'       tlSmf(O) - Spot MG record if found
'
'
    Dim tlWSmf As SMF
    Dim ilRet As Integer
    Dim tlSmfSrchKey2 As LONGKEY0   'SdfCode
    Dim ilSmfRecLen As Integer

    tlSmf.lChfCode = 0
    tlSmf.iOrigSchVef = tlSdfExt.iVefCode
    tlSmf.sMGSource = "A"
    mFindSmf = False
    If (tlSdfExt.sSchStatus = "G") Or (tlSdfExt.sSchStatus = "O") Then
        
        ilSmfRecLen = Len(tlWSmf)
        tlSmfSrchKey2.lCode = tlSdfExt.lCode
        ilRet = btrGetEqual(hlSmf, tlWSmf, ilSmfRecLen, tlSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            tlSmf = tlWSmf 'return SMF
            mFindSmf = True
            Exit Function
        End If
    End If
End Function
'
'                       mCalcTotalGrossForLines - cycle thru the lines and accumu the total gross for contract
'                       vs getting the $ from the header field which may be incorrect,  also doesnt
'                       include NTR
Public Sub mCalcTotalGrossForLines()
Dim ilClf As Integer

Dim llLineGross(0 To 2) As Long     'index zero ignored
Dim llStartDates(0 To 2) As Long    'index zero ignored
Dim llTotalNTR As Long
Dim llTempGross As Long
Dim slSharePct As String
Dim slGrossAmount As String
Dim llSplitGrossAmt As Long
Dim slStr As String
Dim ilMnfSubCo As Integer
Dim ilLoopOnSlsp As Integer

        lmTotalGross = 0
        llTotalNTR = 0
        llLineGross(1) = 0
        For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1  'loop through all the lines in the contract
            tmClf = tgClf(ilClf).ClfRec
            gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llStartDates(1)
            gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llStartDates(2)
            
            If llStartDates(2) + 1 > llStartDates(1) Then      'Cancel before start order and user doesnt want it shown on report
                gBuildFlights ilClf, llStartDates(), 1, 2, llLineGross(), 1, tgClf(), tgCff()
                'need to build splits here because if using Sub-companies, get the vehicles subcompany to find the slsp split in header
                mCalcSplitSlsp tmClf.iVefCode, llLineGross()
'
'                ReDim llSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
'                ReDim ilSlfCode(0 To 9) As Integer             '4-20-00
'                ReDim llSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)
'
'                ilMnfSubCo = gGetSubCmpy(tgChf, ilSlfCode(), llSlfSplit(), tmClf.iVefCode, False, llSlfSplitRev())
'
'                'Calculate the split revenue
'                llTempGross = llLineGross(1)
'                For ilLoopOnSlsp = LBound(llSlfSplit) To UBound(llSlfSplit)
'                    If llSlfSplit(ilLoopOnSlsp) > 0 Then
'                        slSharePct = gLongToStrDec(llSlfSplit(ilLoopOnSlsp), 4)       'slsp share
'                        slGrossAmount = gLongToStrDec(llLineGross(1), 2)
'                        slStr = gMulStr(slSharePct, slGrossAmount)                       ' gross portion of possible split
'                        llSplitGrossAmt = Val(gRoundStr(slStr, "1", 0))
'                        If llTempGross - llSplitGrossAmt < 0 Then
'                            lmSlspRevShare(ilLoopOnSlsp) = lmSlspRevShare(ilLoopOnSlsp) + llTempGross            'cannot exceed the gross accumulated
'                            Exit For
'                        Else
'                            llTempGross = llTempGross - llSplitGrossAmt
'                            lmSlspRevShare(ilLoopOnSlsp) = lmSlspRevShare(ilLoopOnSlsp) + llSplitGrossAmt
'                        End If
'                    End If
'                Next ilLoopOnSlsp
'
'
'                lmTotalGross = lmTotalGross + llLineGross(1)       'total $ for line
                llLineGross(1) = 0
            End If
        Next ilClf                                      'loop thru schedule lines
        
        Exit Sub
End Sub
'
'                           mExportNTR - process NTR items if applicable
'                       <input> None
'                       <output> None
'                       Return:  true if Error found
Public Function mExportNTR() As Boolean
Dim blError As Boolean
Dim ilRet As Integer
Dim llAmount(0 To 2) As Long    'index zero ignored
Dim ilLoop As Integer
Dim llDate As Long
Dim ilInx As Integer
Dim slDate As String
Dim slStrAmt As String
Dim ilCommPct As Integer
Dim blExportErr As Boolean
Dim ilAgyInx As Integer

        'find NTR for $
        blError = False
        tmSbfSrchKey0.lChfCode = tgChf.lCode
        tmSbfSrchKey0.iDate(0) = 0
        tmSbfSrchKey0.iDate(1) = 0
        tmSbfSrchKey0.sTranType = " "
        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = tgChf.lCode)
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mExportNTRErr
                gBtrvErrorMsg ilRet, "mExportNTR (GetNext):SBF", ExportSalesForce
                On Error GoTo 0
                blError = True
                mExportNTR = blError
                Exit Function
            End If

            gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
            If (llDate >= lmUserStartDate And llDate <= lmUserEndDate) And tmSbf.sTranType = "I" Then
                  
                  'get vehicle information thats common to all spots in the same vehicle, once per vehicle
                  'but for outside and makegoods, they can be different vehicle than the vehicle of the line it came from,
                  'so retrieve vehicle info
                  ilInx = gBinarySearchVef(tmSbf.iAirVefCode)
                  If ilInx >= 0 Then
                      tmSpotDetail.sVG2SubT = Trim$(tmVGINFO(ilInx).sVG2SubT)
                      tmSpotDetail.sVG3Mkt = Trim$(tmVGINFO(ilInx).sVG3Mkt)
                      tmSpotDetail.sVG4Fmt = Trim$(tmVGINFO(ilInx).sVG4Fmt)
                      tmSpotDetail.sVG5Rsch = Trim$(tmVGINFO(ilInx).sVG5Rsch)
                      tmSpotDetail.sVG6SubC = Trim$(tmVGINFO(ilInx).sVG6SubC)
    
                  Else
                      tmSpotDetail.sVG2SubT = ""
                      tmSpotDetail.sVG3Mkt = ""
                      tmSpotDetail.sVG4Fmt = ""
                      tmSpotDetail.sVG5Rsch = ""
                      tmSpotDetail.sVG6SubC = ""

                End If
                           
                tmSpotDetail.sSpotID = Trim$(Str(tmSbf.lCode))                'internal ntr id
                slDate = Format$(llDate, "m/d/yy")
                tmSpotDetail.sSpotDate = slDate  'spot date

                tmSpotDetail.sSpotTimeAired = ""
                tmSpotDetail.sSpotLength = ""   'spot length
                tmSpotDetail.sAirTimeNTR = "N"
                tmSpotDetail.sNTRType = ""              'ntr type
                For ilLoop = LBound(tmMnfItem) To UBound(tmMnfItem) - 1
                    If tmMnfItem(ilLoop).iCode = tmSbf.iMnfItem Then
                        tmSpotDetail.sNTRType = Trim$(tmMnfItem(ilLoop).sName)
                        Exit For
                    End If
                Next ilLoop
                tmSpotDetail.sVefCode = Trim$(Str(tmSbf.iAirVefCode))
                tmSpotDetail.sContractID = Trim$(Str$(tgChf.lCode))
                llAmount(1) = tmSbf.iNoItems * tmSbf.lGross
                mStripCents llAmount(1), slStrAmt                       'detail ntr record
                tmSpotDetail.sGrossSpotPrice = Trim$(slStrAmt)
                'accumulate NTR items into total gross for contract
                mCalcSplitSlsp tmSbf.iAirVefCode, llAmount()
                
                tmSpotDetail.sAgyCommPct = ""
                If tmSbf.sAgyComm = "Y" Then
                    If tgChf.iAgfCode = 0 Then          'direct
                        ilCommPct = 0                'no commission
                        slStrAmt = ".00"
                    Else
                        ilAgyInx = gBinarySearchAgf(tgChf.iAgfCode)
                        If ilAgyInx >= 0 Then
                            ilCommPct = tgCommAgf(ilAgyInx).iCommPct
                            slStrAmt = Trim$(gLongToStrDec(CLng(ilCommPct), 2))
                        End If
                    End If
                    tmSpotDetail.sAgyCommPct = Trim$(slStrAmt)
                End If
                    
                blExportErr = mWriteExportSpots(hmFileHandle(FILE_SPOTS), FILE_SPOTS)
            ElseIf tmSbf.sTranType = "I" Then              'right contract, but outside dates requested.  Need to total the NTR for total amt of order
                llAmount(1) = (tmSbf.iNoItems * tmSbf.lGross)
                mCalcSplitSlsp tmSbf.iAirVefCode, llAmount()

            End If
            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        
        mExportNTR = blError
        Exit Function
        
mExportNTRErr:
    On Error GoTo 0
    ilRet = Err.Number
    Resume Next
End Function
'
'                   mStripCents - strip cents off the string if .00 cents; otherwise use decimal and cents
'                   no commas string
'                   <input> llAmount - value to be converted to string (w/wo cents)
'                   <output> slStrAmt - amt in string form
'
Public Sub mStripCents(llAmount As Long, slStrAmt As String)
Dim ilRemainder As Integer
Dim slStripCents As String


        ilRemainder = llAmount Mod 100
        If ilRemainder = 0 Then         'strip off the pennies if whole number
            slStripCents = Trim$(gLongToStrDec(llAmount, 2))
            slStrAmt = Trim$(Mid$(slStripCents, 1, Len(slStripCents) - 3))
        Else
           slStrAmt = Trim$(gLongToStrDec(llAmount, 2))
        End If
        Exit Sub
End Sub

Public Sub mCalcSplitSlsp(ilVefCode As Integer, llGross() As Long)
Dim llTempGross As Long
Dim slSharePct As String
Dim slGrossAmount As String
Dim slStr As String
Dim llSplitGrossAmt As Long
Dim ilMnfSubCo As Integer
Dim ilLoopOnSlsp As Integer

            ReDim llSlfSplit(0 To 9) As Long           '4-20-00 slsp slsp share %
            ReDim ilSlfCode(0 To 9) As Integer             '4-20-00
            ReDim llSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

            ilMnfSubCo = gGetSubCmpy(tgChf, ilSlfCode(), llSlfSplit(), ilVefCode, False, llSlfSplitRev())
            
            'Calculate the split revenue
            llTempGross = llGross(1)
            For ilLoopOnSlsp = LBound(llSlfSplit) To UBound(llSlfSplit)
                If llSlfSplit(ilLoopOnSlsp) > 0 Then
                    slSharePct = gLongToStrDec(llSlfSplit(ilLoopOnSlsp), 4)       'slsp share
                    slGrossAmount = gLongToStrDec(llGross(1), 2)
                    slStr = gMulStr(slSharePct, slGrossAmount)                       ' gross portion of possible split
                    llSplitGrossAmt = Val(gRoundStr(slStr, "1", 0))
                    If llTempGross - llSplitGrossAmt < 0 Then
                        lmSlspRevShare(ilLoopOnSlsp) = lmSlspRevShare(ilLoopOnSlsp) + llTempGross            'cannot exceed the gross accumulated
                        Exit For
                    Else
                        llTempGross = llTempGross - llSplitGrossAmt
                        lmSlspRevShare(ilLoopOnSlsp) = lmSlspRevShare(ilLoopOnSlsp) + llSplitGrossAmt
                    End If
                End If
            Next ilLoopOnSlsp

            lmTotalGross = lmTotalGross + llGross(1)       'total $ for line
        Exit Sub
End Sub

Public Sub mInsertOutFileNameMsg(blErrorOK As Boolean, ilWhichFile As Integer)
        
Dim ilSelectFile As Integer
Dim slFileStatus As String
Dim slFile As String

        ilSelectFile = ilWhichFile
        If blErrorOK Then
            slFileStatus = "Completed- "
        Else
            slFileStatus = "Terminated- "
        End If
        
        Select Case ilSelectFile
            Case FILE_PRODPROT
                slFileStatus = "Product Protection " & slFileStatus
            Case FILE_VEHICLE
                slFileStatus = "Vehicle " & slFileStatus
            Case FILE_SLSP
                slFileStatus = "Salespeople " & slFileStatus
            Case FILE_SS
                slFileStatus = "Sales Source " & slFileStatus
            Case FILE_OFFICE
                slFileStatus = "Sales Office " & slFileStatus
            Case FILE_ADV
                slFileStatus = "Advertiser " & slFileStatus
            Case FILE_AGY
                slFileStatus = "Agency " & slFileStatus
                Case FILE_CONTRACT
                slFileStatus = "Contract " & slFileStatus
            Case FILE_SPOTS
                 slFileStatus = "Spots " & slFileStatus
       End Select
        
        gLogMsg slFileStatus & " Saved in " & smExportName(ilWhichFile), "ExportSalesForce.txt", False
        lbcInfo(0).AddItem slFileStatus & " Saved in " & smExportName(ilWhichFile)
        
        Exit Sub
        
End Sub
'
'this is required to intiate the pvsw engine; otherwise err 3012 occurs with spfread
'
Private Sub mTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    gInitGlobalVar                  'this is called twice (one here, and gInitStandAlone routines), but required here before initiating the pvsw engine to get all the .ini values for paths
                                    'Error 3012 occurs with gspfread if not initiated
    hlSpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        Sleep 2000
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    ilRecLen = Len(tlSpf)
    ilRet = btrGetFirst(hlSpf, tlSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        Sleep 2000
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    btrDestroy hlSpf
End Sub

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "SFE"
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



