VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Begin VB.Form ExportInvSpots 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Inventory & Spots"
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
   Icon            =   "ExportInvSpots.frx":0000
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
      Picture         =   "ExportInvSpots.frx":08CA
      ScaleHeight     =   6015
      ScaleWidth      =   7485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7545
      Begin VB.ListBox lbcInfo 
         Height          =   900
         Index           =   1
         ItemData        =   "ExportInvSpots.frx":A217
         Left            =   570
         List            =   "ExportInvSpots.frx":A219
         TabIndex        =   24
         Top             =   2790
         Width           =   6255
      End
      Begin VB.ListBox lbcInfo 
         Height          =   690
         Index           =   0
         ItemData        =   "ExportInvSpots.frx":A21B
         Left            =   570
         List            =   "ExportInvSpots.frx":A21D
         TabIndex        =   23
         Top             =   3885
         Width           =   6255
      End
      Begin VB.CommandButton cmcTo 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
         Height          =   285
         Index           =   1
         Left            =   5670
         TabIndex        =   21
         Top             =   3465
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox plcTo 
         Height          =   375
         Index           =   1
         Left            =   1170
         ScaleHeight     =   315
         ScaleWidth      =   4245
         TabIndex        =   19
         Top             =   3375
         Visible         =   0   'False
         Width           =   4305
         Begin VB.TextBox edcTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   15
            TabIndex        =   20
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
         Begin VB.TextBox edcTo 
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
         Left            =   5670
         TabIndex        =   8
         Top             =   3015
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
         Top             =   2160
         Width           =   1200
      End
      Begin VB.ListBox lbcVehicle 
         Appearance      =   0  'Flat
         Height          =   1500
         ItemData        =   "ExportInvSpots.frx":A21F
         Left            =   2985
         List            =   "ExportInvSpots.frx":A221
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   1140
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CheckBox ckcAll 
         BackColor       =   &H00FFFFC0&
         Caption         =   "All Vehicles"
         Height          =   210
         Left            =   2985
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
         Left            =   1410
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
      Begin VB.Label lacSaveIn 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Save Spots"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   225
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   570
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   2925
         Visible         =   0   'False
         Width           =   6255
      End
      Begin DZACTXLibCtl.dzactxctrl zpcDZip 
         Left            =   705
         OleObjectBlob   =   "ExportInvSpots.frx":A223
         Top             =   4710
      End
      Begin VB.Label lacSaveIn 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Save Inv"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   16
         Top             =   3075
         Visible         =   0   'False
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
Attribute VB_Name = "ExportInvSpots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExportInvSpots.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text

Dim hmInv As Integer            'Inv file handle
Dim hmSpots As Integer          'spots file handle

Dim smExportNameInv As String
Dim smExportNameSpots As String
Dim imFirstActivate As Integer
Dim lmCntrNo As Long    'for debugging purposes to filter a single contract
Dim lmCntrCode As Long  'contract code so header doesnt have to be read to match all spots

Dim imSetAll As Integer
Dim imAllClicked As Integer

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
Dim tmSdfSrchKey4 As SDFKEY4            'date, chfcode

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

Dim hmSlf As Integer
Dim tmSlf As SLF                  'Salesperson  record image
Dim imSlfRecLen As Integer        'Slsp record length

Dim hmSof As Integer
Dim tmSof As SOF                  'Slsp office  record image
Dim imSofRecLen As Integer        'SOF record length
Dim tmSofArray() As SOF

Dim hmAdf As Integer
Dim tmAdf As ADF                  'Advt  record image
Dim imAdfRecLen As Integer        'ADF record length

Dim hmAgf As Integer
Dim tmAgf As AGF                  'AGency  record image
Dim imAgfRecLen As Integer        'AGF record length

Dim hmMnf As Integer
Dim tmMnf As MNF                  'Multiname  record image
Dim tmMnfVGroup() As MNF
Dim tmMnfCompCode() As MNF
Dim imMnfRecLen As Integer        'MNF record length


Dim hmAnf As Integer
Dim tmAnf As ANF                  'Named avail record image
Dim tmAnfArray() As ANF
Dim imAnfRecLen As Integer        'ANF record length


Dim hmLcf As Integer
Dim tmLcf As LCF                'Log calendar  record image
Dim imLcfRecLen As Integer        'Log calendar record length

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim lmNowDate As Long
Dim lmUserStartDate As Long         'user entered start date
Dim lmUserEndDate As Long           'user entered end date
Dim lmLatestSDFDate As Long         'Latest sdf spot date (for tfn)
Dim imAutoRun As Integer            '1 = auto run flag, 0 = manual
Dim imClosed As Integer

Dim imVehiclesInPast() As Integer        'array of vehicles to process in the past
Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT

Private Type SPOTHEADER
    sContractType As String * 1
    sAdvtName As String * 30
    sProduct As String * 35
    sAgyName As String * 40
    sContract As String * 9
    sPrimComp As String
    sSecComp As String
    sSlspFirstName As String
    sSlspLastName As String
    sSlspSplitPct As String
    sOffice As String
    sTradePct As String * 3
End Type

Private Type SPOTVEFINFO
    sVehicleName As String
    sVehicleCode As String
    sVGGroup1 As String
    sVGGroup2 As String
    sVGGroup3 As String
    sVGGroup4 As String
    sVGGroup5 As String
    End Type


Private Type SPOTDETAIL
    sSpotDate As String * 8
    sSpotDay As String * 3
    sSpotTimeSch As String * 8
    sSpotTimeAired As String * 8
    sSpotStatus As String * 1
    sSpotType As String * 1
    sGame As String * 3
    sSpotLength As String * 3
    sGrossSpotPrice As String
    sAgyCommAmt As String
    sNetAmt As String
    sCopyFlag As String * 1
End Type
    
Private Type CHFFOUND
    sKey As String * 10
    lCode As Long
    lCntrNo As Long
End Type
 
Dim tmSpotDetail() As SPOTDETAIL
Dim tmSpotHeader As SPOTHEADER
Dim tmSpotVefInfo As SPOTVEFINFO
Dim tmChfFound() As CHFFOUND

Private Const DELIMITER = "|"         '| vertical bar, or pipe character


Private Sub edcTo_GotFocus(ilIndex As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
'    gShowBranner
End Sub



Private Sub Form_Load()
    If App.PrevInstance Then
        If imAutoRun = 0 Then
            MsgBox "Only one copy of ExportInvSpots can be run at a time ", vbInformation + vbOKOnly, "Counterpoint"
        End If
        gLogMsg "Second copy of ExportInvSpots path: " & App.Path & " from " & Trim$(gGetComputerName()), "ExportInvSpots.txt", False
        End
    End If
    gSetBkgdMode
    'igBkgdProg = 6
    mInit
    If imAutoRun = 0 Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMinimized
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
    '    ilRet = MsgBox("Stop the Export Inventory & Spots", vbQuestion + vbYesNo, "Stop Service")
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
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    btrStopAppl
    Set ExportInvSpots = Nothing   'Remove data segment
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
'        gLogMsg "Current Mode: UserInput.", "ExportInvSpots.txt", False
    Else
        igBkgdProg = 12
'        gLogMsg "Current Mode: Background.", "ExportInvSpots.txt", False
    End If
    gInitStdAlone ExportInvSpots, slStr, ilTestSystem
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "ExportInvSpots.Frm"
    
    If igBkgdProg = 0 Then
        gLogMsg "Current Mode: UserInput.", "ExportInvSpots.txt", False
    Else
        gLogMsg "Current Mode: Background.", "ExportInvSpots.txt", False
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
    'Dim slDateTime As String
    Dim ilOk As Integer
    
    If imAutoRun = 0 Then           'manual mode
        'check for validity of date in case it was changed
        
        slDate = edcStart.Text
        If Not gValidDate(slDate) Then
            edcStart.SetFocus
            Exit Sub
        End If
        lmUserStartDate = gDateValue(slDate)
        
        slDate = edcEnd.Text
        If Not gValidDate(slDate) Then
            edcEnd.SetFocus
            Exit Sub
        End If
        
        'test start date is prior to end date
        
        lmUserEndDate = gDateValue(slDate)
        
        slTodayDate = Format$(lmUserStartDate, "m/d/yy")
        
        If lmUserStartDate > lmUserEndDate Then
            edcEnd.SetFocus
            Exit Sub
        End If
        
        gObtainYearMonthDayStr slTodayDate, True, slYear, slMonth, slDay
        edcStart.Text = slTodayDate
    
        slLatestDate = Format$(lmUserEndDate, "m/d/yy")
        gObtainYearMonthDayStr slLatestDate, True, slLatestYear, slLatestMonth, slLatestDay
        slDefaultFileName = "Inv-" & Trim$(slMonth) & Trim$(slDay) & Trim$(Mid$(slYear, 3)) & "-" & Trim$(slLatestMonth) & Trim$(slLatestDay) & Trim$(Mid$(slLatestYear, 3))
        slDefaultSpotFileName = "Spot-" & Trim$(slMonth) & Trim$(slDay) & Trim$(Mid$(slYear, 3)) & "-" & Trim$(slLatestMonth) & Trim$(slLatestDay) & Trim$(Mid$(slLatestYear, 3))
        
        ilRet = 0
        'On Error GoTo lacMinErr:
        smExportNameInv = sgExportPath & slDefaultFileName & ".txt"
        'slDateTime = FileDateTime(smExportNameInv)
        ilRet = gFileExist(smExportNameInv)
        smExportNameSpots = sgExportPath & slDefaultSpotFileName & ".txt"
        'slDateTime = FileDateTime(smExportNameSpots)
        ilRet = gFileExist(smExportNameSpots)

        edcTo(0).Text = smExportNameInv
        edcTo(0).Visible = True
        edcTo(1).Text = smExportNameSpots
        edcTo(1).Visible = True
    

        
        If Not imExporting Then
            Screen.MousePointer = vbHourglass
            ilOk = mExportAll()
            Screen.MousePointer = vbDefault
            If ilOk Then
                lacStop.Caption = "Done"
                lacMin.Enabled = False
            End If
        End If
    Else
        Me.WindowState = vbMinimized
    End If
    Exit Sub
'lacMinErr:
'    ilRet = 1
'    Resume Next
End Sub
Private Sub lacStop_Click()
    If Not imExporting Then
        imTerminate = True
        mTerminate
    Else
        imTerminate = True
        Close #hmInv
        Close #hmSpots
        mCloseInvSpotFiles
        Screen.MousePointer = vbDefault
        imExporting = False
        lbcInfo(1).AddItem "Export Cancelled "
        lacStop.Caption = "Done"
        gLogMsg "Export Cancelled", "ExportInvSpots.txt", False
        mTerminate
    End If
End Sub


Private Sub mBtrErrorMsg(slMsg As String, ilInRet As Integer)
    Dim ilRet As Integer
    
    ilRet = ilInRet
    If ilRet >= 30000 Then
        ilRet = csiHandleValue(0, 7)
    End If
    gLogMsg slMsg & Trim$(Str$(ilRet)), "ExportInvSpots.txt", False
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
    ilRet = gPopUserVehicleBox(ExportInvSpots, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + ACTIVEVEH + VEHSPORT, lbcVehicle, tgUserVehicle(), sgUserVehicleTag)

    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExportInvSpots
        On Error GoTo 0
    End If
    
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'
'                   mCreateInv - create inventory for default dates (today minus 3 days to end of spot file), or user entered start/end dates
'                   Process for all conventional, selling and game vehicles; one vehicle at a time for all dates indicated.
'
Function mCreateInv() As Integer
    Dim ilRet As Integer
    Dim ilError As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    Dim llEarliestLcfDate As Long
    Dim llLatestLcfDate As Long
    Dim llDate As Long
    Dim ilDate(0 To 1) As Integer
    Dim llTime As Long
    Dim slDate As String
    Dim ilType As Integer

    ilError = 0             'assume everything is OK
    ilIndex = gBinarySearchVpf(tmVef.iCode)
    If ilIndex <> -1 Then
        
        'show Processing vehicle & dates on caption screen & file
        gLogMsg "Processing Inventory " & Trim$(tmVef.sName) & " for " & Format(lmUserStartDate, "m/d/yy") & " - " & Format(lmUserEndDate, "m/d/yy"), "ExportInvSpots.txt", False
        lbcInfo(1).AddItem "Processing " & Trim$(tmVef.sName) & " for " & Format(lmUserStartDate, "m/d/yy") & " - " & Format(lmUserEndDate, "m/d/yy")
        lbcInfo(1).Visible = True
        ilType = 0
        'use the ssf for inventory only
        For llDate = lmUserStartDate To lmUserEndDate
            DoEvents
            
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
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilLoop)
                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                        ilRet = mWriteExportInv()
                        If Not ilRet Then
                            ilError = 2
                            Exit For
                        End If
                    End If
                Next ilLoop
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tmVef.sType = "G" Then
                    ilType = tmSsf.iType
                End If
            Loop
            slDate = Format$(llDate, "m/d/yy")
            
        Next llDate
    Else                    'didnt find vehicle options
        ilError = 1
    End If
    
    mCreateInv = ilError
End Function

'
'
'
'           mCloseInvSpotFiles - Close all applicable files for
'                       projection Export
'
Sub mCloseInvSpotFiles()
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
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmAnf)
    
    btrDestroy hmVef
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmSdf
    btrDestroy hmSsf
    btrDestroy hmLcf
    btrDestroy hmVsf
    btrDestroy hmSmf
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmSof
    btrDestroy hmAdf
    btrDestroy hmAgf
    btrDestroy hmMnf
    btrDestroy hmAnf
End Sub

'
'
'           mOpenFiles - open files applicable to Export Inventory and Spots
'
'
Function mOpenInvSpotFiles() As Integer
    Dim ilRet As Integer
    Dim ilTemp As Integer
    Dim ilError As Integer
    Dim slStamp As String
    Dim slMnfStamp As String

    ilError = False

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen VEF)", ExportInvSpots
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen CHF)", ExportInvSpots
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen CLF)", ExportInvSpots
    On Error GoTo 0
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen CFF)", ExportInvSpots
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    
    'opened in mInit to get latest spot date
'    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mOpenInvSpotFilesErr
'    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen SDF)", ExportInvSpots
'    On Error GoTo 0
'    imSdfRecLen = Len(tmSdf)
    
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen SSF)", ExportInvSpots
    On Error GoTo 0
    imSsfRecLen = Len(tmSsf)
    
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen LCF)", ExportInvSpots
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)

    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen SMF)", ExportInvSpots
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)
    
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen VSF)", ExportInvSpots
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen SLF)", ExportInvSpots
    On Error GoTo 0
    imSlfRecLen = Len(tmSlf)

    
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen SOF)", ExportInvSpots
    On Error GoTo 0
    imSofRecLen = Len(tmSof)

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen ADF)", ExportInvSpots
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen AGF)", ExportInvSpots
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen MNF)", ExportInvSpots
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mOpenInvSpotFilesErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen ANF)", ExportInvSpots
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)


    gObtainAdvt
    gObtainAgency
    gObtainVef
    ilRet = gObtainSalesperson()
    
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

    gPopAnf hmAnf, tmAnfArray()              'populate the Named Avails
    mPopSof hmSof, tmSofArray()             'populate the Sales Offices
    
    mOpenInvSpotFiles = ilError
    Exit Function

mOpenInvSpotFilesErr:
    ilError = True
    Return
End Function

'
'                       mWriteExportInv - write a record for each day of all conventional, selling & event vehicles
'                       for as many days as requested.
'                   Create file Inv-mmddyy-mmddyy.txt (pipe character delimited)
'Field   Type    Max Format and Example
'1.  Vehicle A   40  Raw Dog
'2.  Vehicle ID  N   10  Internal CSI #
'3.  Game #  N   3   001
'4.  Avail Date  A-N 8   mm/dd/yy, 1/1/11
'5.  Avail Day   A   3   MON, TUE, WED, THU, FRI, SAT, or SUN
'6.  Avail Time  A-N 8   hh:mm:ss  military time
'7.  Avail Name  A   20  Weather, etc.
'8.  Max Units   N   3   2
'9.  Max Seconds N   3   60

'
Private Function mWriteExportInv() As Integer
    Dim ilLoop As Integer
    Dim slRecord As String
    Dim ilIndex As Integer
    Dim slVehicle As String
    Dim ilError As Integer
    Dim ilRecdLen As Integer
    Dim slMonth As String
    Dim slYear As String
    Dim slStr As String
    Dim llAvgRate As Long
    Dim ilRet As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim slDayDescr As String * 3
    Dim slDayOfWeek As String * 21
    Dim slAvailTime As String
    Dim slAvailName As String
    Dim slDate As String
    
    DoEvents
    slDayOfWeek = "MONTUEWEDTHUFRISATSUN"
    ilError = False
   
    'format the month info for a contract/vehicle
    slVehicle = ""
    ilIndex = gBinarySearchVef(tmVef.iCode)
    If ilIndex <> -1 Then
        slVehicle = Trim$(tmVef.sName)
        gUnpackDate tmSsf.iDate(0), tmSsf.iDate(1), slDate
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        'obtain the day of the week
        ilDay = gWeekDayStr(slDate)
        slDayDescr = Mid$(slDayOfWeek, (ilDay * 3) + 1, 3)
        DoEvents
        
        slRecord = Trim$(tmVef.sName) & DELIMITER               'vehicle name & vertical bar
        slRecord = slRecord & Trim$(Str$(tmVef.iCode)) & DELIMITER
        slRecord = slRecord & Trim$(Str$(tmSsf.iType)) & DELIMITER   'game #, if applicable
        slRecord = slRecord & Trim$(slDate) & DELIMITER        'Inventory date
        slRecord = slRecord & slDayDescr & DELIMITER           'day of week
        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "M", "1", slAvailTime         'time of avail in military time
        slRecord = slRecord & Trim$(Mid$(slAvailTime, 1, 2)) & ":" & Trim$(Mid$(slAvailTime, 3, 5)) & DELIMITER
        
        slAvailName = ""
        For ilLoop = LBound(tmAnfArray) To UBound(tmAnfArray) - 1       'search of avail name description
            If tmAvail.ianfCode = tmAnfArray(ilLoop).iCode Then
                slAvailName = Trim$(tmAnfArray(ilLoop).sName)
                Exit For
            End If
        Next ilLoop
        slRecord = slRecord & Trim$(slAvailName) & DELIMITER

        slRecord = slRecord & Trim$(Str$(tmAvail.iAvInfo And &HF)) & DELIMITER
        slRecord = slRecord & Trim$(Str$(tmAvail.iLen))
        

        ilRecdLen = Len(Trim(slRecord))
        ilRet = 0
        On Error GoTo mWriteExportInvErr
        Print #hmInv, Left(slRecord, ilRecdLen)
        On Error GoTo 0
        If ilRet <> 0 Then
            imExporting = False
            gLogMsg "Error writing to Export Inv File, Error # " & Str$(ilRet), "ExportInvSpots.txt", False
            mWriteExportInv = False
            Exit Function
        End If
  

    End If

    mWriteExportInv = True
    Exit Function

mWriteExportInvErr:
    ilRet = Err.Number
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
Private Sub cmcTo_Click(ilIndex As Integer)
    CMDialogBox.DialogTitle = "Export To File"
    CMDialogBox.Filter = "Comma|*.CSV|ASC|*.Asc|Text|*.Txt|All|*.*"
    CMDialogBox.InitDir = Left$(sgExportPath, Len(sgExportPath) - 1)
    CMDialogBox.DefaultExt = ".Csv"
    CMDialogBox.flags = cdlOFNCreatePrompt
    CMDialogBox.Action = 1 'Open dialog
    edcTo(0).Text = CMDialogBox.fileName
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
    If ilIndex = 0 Then
        If edcTo(0).Text = "" Then
            edcTo(0).Text = smExportNameInv
        End If
    Else
        If edcTo(1).Text = "" Then
            edcTo(1).Text = smExportNameSpots
        End If
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
    ckcAll.Value = vbChecked
    
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mOpenInvSpotFiles (btrOpen SDF)", ExportInvSpots
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Traffic
    Unload ExportInvSpots
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
    'Dim slDateTimeSpot As String
    Dim slDefaultFileName As String
    Dim slDefaultSpotFileName As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slTodayDate As String
    Dim slLatestSDFDate As String
    Dim slLatestMonth As String
    Dim slLatestDay As String
    Dim slLatestYear As String
    Dim ilOk As Integer
    
        tmcClick.Enabled = False
        If imTerminate Then
            lacStop_Click
            Exit Sub
        End If
        
        slTodayDate = Format$(lmNowDate - 3, "m/d/yy")
        gObtainYearMonthDayStr slTodayDate, True, slYear, slMonth, slDay
        edcStart.Text = slTodayDate
      
        slLatestSDFDate = Format$(lmLatestSDFDate, "m/d/yy")
        gObtainYearMonthDayStr slLatestSDFDate, True, slLatestYear, slLatestMonth, slLatestDay
        'Determine name of export for inventory & spots (.txt file)
        'for now, do not create unique names that append a new letter of the alphabet.
        'Always ask to overwrite
        slRepeat = ""
        slDefaultFileName = "Inv-" & Trim$(slMonth) & Trim$(slDay) & Trim$(Mid$(slYear, 3)) & "-" & Trim$(slLatestMonth) & Trim$(slLatestDay) & Trim$(Mid$(slLatestYear, 3))
        slDefaultSpotFileName = "Spot-" & Trim$(slMonth) & Trim$(slDay) & Trim$(Mid$(slYear, 3)) & "-" & Trim$(slLatestMonth) & Trim$(slLatestDay) & Trim$(Mid$(slLatestYear, 3))
        
        lmUserStartDate = gDateValue(slTodayDate)
        lmUserEndDate = lmLatestSDFDate
        
        ilRet = 0
        'On Error GoTo cmcExportDupNameErr:
        smExportNameInv = sgExportPath & slDefaultFileName & ".txt"
        'slDateTime = FileDateTime(smExportNameInv)
        ilRet = gFileExist(smExportNameInv)
    
        smExportNameSpots = sgExportPath & slDefaultSpotFileName & ".txt"
        'slDateTimeSpot = FileDateTime(smExportNameSpots)
        ilRet = gFileExist(smExportNameSpots)
    
        edcTo(0).Text = smExportNameInv
        edcTo(0).Visible = True
        edcTo(1).Text = smExportNameSpots
        edcTo(1).Visible = True
        
    
        
        If imAutoRun = 1 Then         '9-07-05 if auto run, activate the task
            ilOk = mExportAll()
            Screen.MousePointer = vbDefault
'            If ilOK Then           'auto run does not show the results
'                lacMin.Caption = "Done"
'                lacMin.Enabled = False
'            End If
        End If
    Exit Sub
'cmcExportDupNameErr:
'    ilRet = 1
'    Resume Next
End Sub

Public Sub mSetCommands()

    If lbcVehicle.SelCount > 0 And edcTo(0).Text <> "" And edcTo(1).Text <> "" Then
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
'                   ExportInvSpots.txt is a message file placed under data\messages folder
'
Private Function mExportAll()
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
    
    'lbcInfo(0).Visible = False

    If imExporting Then
        Exit Function
    End If


    lmCntrNo = 0                'ths is for debugging on a single contract
    slStr = ExportInvSpots!edcContract
    If slStr <> "" Then
        lmCntrNo = Val(slStr)
    End If

    
    smExportNameInv = Trim$(edcTo(0).Text)
    smExportNameSpots = Trim$(edcTo(1).Text)
    
    If imAutoRun = 0 Then           'manual, disable the export button
        lacMin.Enabled = False
    End If
    
    gLogMsg "** Export Inventory & Spots, begin: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "ExportInvSpots.txt", False
    
    For ilLoop = 0 To 1
        If ilLoop = 0 Then
            slExportName = smExportNameInv
        Else
            slExportName = smExportNameSpots
        End If

        If (InStr(slExportName, ":") = 0) And (Left$(slExportName, 2) <> "\\") Then     'test for absence of colon and not using \\
            slExportName = sgExportPath & slExportName
        End If
    
        ilRet = 0
        'On Error GoTo mExportErr:
        'slDateTime = FileDateTime(slExportName)
        ilRet = gFileExist(slExportName)
        If ilRet = 0 Then
            Kill slExportName
            gLogMsg "** Removing file " & slExportName, "ExportInvSpots.txt", False
        End If
     
        'smExportNameInv = slExportName
        hmExportFile = FreeFile
        
        ilRet = 0
        'Open slExportName For Output As hmExportFile
        ilRet = gFileOpen(slExportName, "Output", hmExportFile)
        If ilRet <> 0 Then
            
            gLogMsg "Open File Error #" & Str$(Err.Number) & slExportName, "ExportInvSpots.txt", False
            Close #hmExportFile
            imExporting = False
            Screen.MousePointer = vbDefault
            If imAutoRun = 0 Then
                MsgBox "Open File Error #" & Str$(Err.Number) & slExportName, vbOKOnly, "Open Error"
            End If
            mExportAll = False          'error
            Exit Function
        End If
        On Error GoTo 0
        If ilLoop = 0 Then
            hmInv = hmExportFile
        Else
            hmSpots = hmExportFile
        End If
    Next ilLoop

    Screen.MousePointer = vbHourglass
    imExporting = True
    
    ilAnyErrors = False         'general error flag
    
    If mOpenInvSpotFiles() = 0 Then
        ilErrorStatus = 0
        If lmCntrNo <> 0 Then
            tmChfSrchKey1.lCntrNo = lmCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_END_OF_FILE Or tmChf.lCntrNo <> lmCntrNo Then
                'exit, no contract found
                lbcInfo(0).AddItem "Contract #" & Trim$(ExportInvSpots!edcContract) & " does not exist, application aborted"
                gLogMsg "Contract #" & Trim$(ExportInvSpots!edcContract) & " does not exist, application aborted", "ExportInvSpots.txt", False
                Close #hmInv
                Close #hmSpots
                mCloseInvSpotFiles
                
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
        
        gLogMsg "** Exporting Inventory to: " & smExportNameInv, "ExportInvSpots.txt", False

        'Process Inventory by vehicle
        ilAnyInvErrors = False
        For ilVehicle = 0 To lbcVehicle.ListCount - 1
            DoEvents
            If lbcVehicle.Selected(ilVehicle) = True Then
                slNameCode = tgUserVehicle(ilVehicle).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVefIndex = gBinarySearchVef(ilVefCode)
                tmVef = tgMVef(ilVefIndex)
                DoEvents
                
                ilError = mCreateInv()
                
                If ilError = 1 Then   'returned with error?
                    lbcInfo(0).AddItem Trim$(tmVef.sName) & " Options table missing"
                    gLogMsg " Export error - " & Trim$(tmVef.sName) & " options table missing", "ExportInvSpots.txt", False
                    ilErrorStatus = 1
                    ilAnyInvErrors = True
                ElseIf ilError = 2 Then
                    ilErrorStatus = ilError
                    ilAnyInvErrors = True
                    Exit For
                End If
            End If
        Next ilVehicle
       
        Close #hmInv
        If ilAnyInvErrors Then     'did all vehicles get processed for inventory without any errors
            lbcInfo(0).AddItem "Inventory Partially Completed, saved in " & smExportNameInv
            gLogMsg "** Export Inventory Partially completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "ExportInvSpots.txt", False
            gLogMsg "** Export has been saved in " & smExportNameInv, "ExportInvSpots.txt", False
        Else
            lbcInfo(0).AddItem "Inventory Completed, saved in " & smExportNameInv
            gLogMsg "** Export Inventory Completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "ExportInvSpots.txt", False
            gLogMsg "** Export Inventory saved in " & smExportNameInv, "ExportInvSpots.txt", False
        End If
        
        ilAnySpotErrors = mCreateSpots()
        Close #hmSpots
        
        If ilAnySpotErrors Then     'did all vehicles get processed for inventory without any errors
            lbcInfo(0).AddItem "Spots Partially Completed, saved in " & smExportNameSpots
            gLogMsg "** Export Spots Partially completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "ExportInvSpots.txt", False
            gLogMsg "** Export Spots saved in " & smExportNameSpots, "ExportInvSpots.txt", False
        Else
            lbcInfo(0).AddItem "Spots Completed, saved in " & smExportNameSpots
            gLogMsg "** Export Spots Completed " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "ExportInvSpots.txt", False
            gLogMsg "** Export Spots saved in " & smExportNameSpots, "ExportInvSpots.txt", False
        End If

        mCloseInvSpotFiles
        Screen.MousePointer = vbDefault

    Else
        lbcInfo(0).AddItem "Open File Error: Export Failed"
        gLogMsg "** Export File Open error " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **", "ExportInvSpots.txt", False
        
    End If
    lbcInfo(0).Visible = True
    Screen.MousePointer = vbDefault
    imExporting = False
    
    
    Erase tmMnfVGroup
    Erase tmMnfCompCode
    Erase tmAnfArray
    
    If imAutoRun = 1 Then     '9-7-05 if auto run, task finished, unload and return to caller
        lacStop_Click
    End If
    mExportAll = True
    If (ilAnyErrors) Or (ilAnyInvErrors) Or (ilAnySpotErrors) Then          'any errors encountered
        mExportAll = False
    End If
        
    Exit Function
'mExportErr:
'    ilRet = Err.Number
'    Resume Next

                              
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
'               mCreateSpots - Obtain all contracts active from user selected start/end dates, or the default dates (today minus 3 days to end of spot file).
'               Go back 1 month from start date to catch makegoods for contracts expired.
'               Process one contract at a time & gather the spots by contract.
'               Create an export entry for every spot.
'
Private Function mCreateSpots() As Integer
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
Dim ilAnySpotErrors As Integer
Dim ilCurrentRecd As Integer
Dim ilInx As Integer
Dim llUpper As Long
Dim ilPrevVefCode As Integer
Dim slDate As String
Dim slDayDescr As String * 3
Dim slDayOfWeek As String * 21
Dim ilDay As Integer
Dim slSpotTime As String
Dim llRecPos As Long
Dim slTime As String
Dim llAmount As Long
Dim ilAgfInx As Integer
Dim ilCommPct As Integer
Dim llNet As Long
Dim llComm As Long
Dim slAmount As String
Dim slSharePct As String
Dim slStr As String
Dim tlCff As CFF
Dim llDate As Long
Dim llCount As Long
Dim llTodayDate As Long
Dim ilFound As Integer
Dim ilLoopOnPast As Integer
Dim ilReadSdf As Boolean
Dim ilAnyPastSpotErrors As Integer
Dim ilSpotTime(0 To 1) As Integer
Dim ilAllDatesInPast As Integer

        ReDim imVehiclesInPast(0 To 0) As Integer

        slDayOfWeek = "MONTUEWEDTHUFRISATSUN"

        'process spots
        gLogMsg "** Exporting Spots to: " & smExportNameSpots, "ExportInvSpots.txt", False

        If lmCntrCode > 0 Then
            gLogMsg "** Exporting Contract #" & Trim$(Str$(edcContract.Text)), "ExportInvSpots.txt", False
            ReDim tlChfAdvtExt(0 To 1) As CHFADVTEXT
            tlChfAdvtExt(0).lCode = lmCntrCode
        Else
            slStartDate = Format$(lmUserStartDate - 30, "m/d/yy")        'go back 1 month to catch any contracts expired with makegoods scheduled into following month
            slEndDate = Format$(lmUserEndDate, "m/d/yy")
            ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
            ilAnySpotErrors = False
            slCntrStatus = "HOGN"                 'statuses: hold, order, unsch hold, uns order
            slCntrType = "CVTRQ"         'all types: PI, DR, etc.  except PSA(p) and Promo(m)
            ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)
            ilRet = gObtainCntrForDate(ExportInvSpots, slStartDate, slEndDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
        End If
        
        'obtain todays date:  prior to todays date needs to get the scheduled times from ssf avails
        slDate = Format$(gNow(), "m/d/yy")
        llTodayDate = gDateValue(gAdjYear(slDate))  'get todays date
        
        slStartDate = Format$(lmUserStartDate, "m/d/yy")        'these are the actual spot gathering dates
        slEndDate = Format$(lmUserEndDate, "m/d/yy")

        ilAllDatesInPast = False
        If llTodayDate + 1 > lmUserEndDate Then        'all dates in the past
            gLogMsg "Processing spots in the future : None", "ExportInvSpots.txt", False
            lbcInfo(1).AddItem "Processing Spots in the future: None"
            ilAllDatesInPast = True
        Else
        gLogMsg "Processing spots in the future for " & Trim$(Format$(llTodayDate + 1, "m/d/yy")) & " - " & Trim$(slEndDate), "ExportInvSpots.txt", False
        lbcInfo(1).AddItem "Processing Spots in the future for " & Trim$(Format$(llTodayDate, "m/d/yy")) & " - " & Trim$(slEndDate)
        End If


        If lmCntrCode = 0 Then                          'if selective contract (contr code exists), dont bother checking to see if all spots will be processed
        mFindUnReferSpots tlChfAdvtExt()            'Determine which spots in the future are not processed because the contract isnt active during the requested period
                                                    'for example, spots are found in future years and the contract has already expired.  How did the spots get so far in the future?
        End If
        
        For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1        'loop on all contracts from the selected user entered dates, retrieve spots one contract at a time
            If tlChfAdvtExt(ilCurrentRecd).sType <> "S" And tlChfAdvtExt(ilCurrentRecd).sType <> "M" Then       'ignore psa(S) and promo(M)
              'read the entire contract in memory to avoid rereading for each spot
              mInitSpotFields
              llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
              ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())       'read in entire contract
    
              If Not ilRet Then
                  gLogMsg "** Unable to obtain Contract Code " & Trim$(Str$(llContrCode)), "ExportInvSpots.txt", False
                  ilAnySpotErrors = True
              Else
                  'put together the header information.  Gather once per contract
                  llUpper = 0
                  tmSpotHeader.sContract = Trim$(Str$(tgChf.lCntrNo))
                  tmSpotHeader.sContractType = tgChf.sType        'regular, reservation, remnant, DR, PI
                  ilInx = gBinarySearchAdf(tgChf.iAdfCode)
                  If ilInx >= 0 Then
                      tmSpotHeader.sAdvtName = Trim$(tgCommAdf(ilInx).sName)
                  End If
                  tmSpotHeader.sProduct = tgChf.sProduct
                  If tgChf.iAgfCode > 0 Then                  'no internal code indicates direct advt
                      ilAgfInx = gBinarySearchAgf(tgChf.iAgfCode)
                      If ilAgfInx >= 0 Then
                          tmSpotHeader.sAgyName = Trim$(tgCommAgf(ilAgfInx).sName)
                      End If
                  End If
                  tmSpotHeader.sPrimComp = mSearchCompCode(tgChf.iMnfComp(0))         'primary competitive code
                  tmSpotHeader.sSecComp = mSearchCompCode(tgChf.iMnfComp(1))          'secondary competitive code
                  tmSpotHeader.sTradePct = Trim$(Str$(tgChf.iPctTrade))
                  mGetSlspFields
                  
                  'gather all the contracts spots for todays date thru either end of spot file or user requested end date
                  ilRet = gObtainCntrSpot(-1, False, llContrCode, -1, "S", slStartDate, slEndDate, tmSdfExtSort(), tmSdfExt(), 1, False)    'find all contract spots, return by vehicle sort
   
                  tmClf.iLine = -1
                  ilPrevVefCode = -1
                  ilReadSdf = True              'need to read sdf in the routine that gathers the spot info to get all the spot detail
                  For llLoopOnSpots = LBound(tmSdfExtSort) To UBound(tmSdfExtSort) - 1
                      llSpotIndex = tmSdfExtSort(llLoopOnSpots).lSdfExtIndex
                      gUnpackDateLong tmSdfExt(llSpotIndex).iDate(0), tmSdfExt(llSpotIndex).iDate(1), llDate
                      'If (llDate >= llTodayDate And llDate <= lmUserEndDate) Or (tmSdfExt(llSpotIndex).sSchStatus = "M" Or tmSdfExt(llSpotIndex).sSchStatus = "C") Then          'process the future spots (todays date thru end of period),
                      If (llDate > llTodayDate And llDate <= lmUserEndDate) Or (tmSdfExt(llSpotIndex).sSchStatus = "M" Or tmSdfExt(llSpotIndex).sSchStatus = "C") Then          'process the future spots (todays date +1 thru end of period),
                                                                                'and all misses and cancelled for the past since theyare not in SSF and wont be found
                        If tmSdfExt(llSpotIndex).iLineNo = tmClf.iLine And tmSdfExt(llSpotIndex).iVefCode = tmClf.iVefCode Then         'same schedule line, determine if n/c, rates, get time dte
                            ilRet = ilRet
                        Else
                            For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                                tmClf = tgClf(ilClf).ClfRec
                                If tmSdfExt(llSpotIndex).iLineNo = tmClf.iLine Then
                                    'if processing a different air vehicle or its a mg or outside, need to ensure to get the air vehicles info
                                    If tmClf.iVefCode <> ilPrevVefCode Or (tmSdfExt(llSpotIndex).sSchStatus = "G" Or tmSdfExt(llSpotIndex).sSchStatus = "O") Then
                                    'If tmSdfExt(llSpotIndex).iVefCode <> ilPrevVefCode Then  'the spot retrieved has the vehicle of the aired spot (if mg or outside), not the vehicle it came from
                                        If ilPrevVefCode <> -1 Then                   '-1 is first time thru, nothing to create yet; otherwise, change in vehicle so write out the previous spot
                                            mWriteExportSpots
                                            ReDim tmSpotDetail(0 To 0) As SPOTDETAIL
                                            llUpper = 0
                                        End If
                                        'get vehicle information thats common to all spots in the same vehicle, once per vehicle
                                        'but for outside and makegoods, they can be different vehicle than the vehicle of the line it came from,
                                        'so retrieve vehicle info
                                        ilPrevVefCode = tmSdfExt(llSpotIndex).iVefCode
                                        ilInx = gBinarySearchVef(tmSdfExt(llSpotIndex).iVefCode)
                                        If ilInx >= 0 Then
                                            tmSpotVefInfo.sVehicleName = Trim$(tgMVef(ilInx).sName)
                                            tmSpotVefInfo.sVehicleCode = Trim$(Str(tgMVef(ilInx).iCode))        'vehicle internal code
                                            tmSpotVefInfo.sVGGroup1 = mGetVGName(tgMVef(ilInx).iMnfVehGp2)            'subtotals
                                            tmSpotVefInfo.sVGGroup2 = mGetVGName(tgMVef(ilInx).iMnfVehGp3Mkt)            'markets
                                            tmSpotVefInfo.sVGGroup3 = mGetVGName(tgMVef(ilInx).iMnfVehGp4Fmt)            'formats
                                            tmSpotVefInfo.sVGGroup4 = mGetVGName(tgMVef(ilInx).iMnfVehGp5Rsch)            'research
                                            tmSpotVefInfo.sVGGroup5 = mGetVGName(tgMVef(ilInx).iMnfVehGp6Sub)            'subcompanies
                                            Exit For
                                        Else
                                            tmSpotVefInfo.sVehicleName = "Vehicle code not found " & Trim$(Str$(ilPrevVefCode))
                                        End If
                                    Else
                                        Exit For            'same vehicle
                                    End If
                                End If
                            Next ilClf
                        End If                  'tmsdfext(llspotindex).ilineno = tmclf.iline
                      
                        
                        'Gather spot detail
'                        gUnpackDate tmSdfExt(llSpotIndex).iDate(0), tmSdfExt(llSpotIndex).iDate(1), slDate
'                        tmSpotDetail(llUpper).sSpotDate = slDate  'spot date
'                        ilDay = gWeekDayStr(slDate)
'                        tmSpotDetail(llUpper).sSpotDay = Mid$(slDayOfWeek, (ilDay * 3) + 1, 3)      'spot date
'                        gUnpackTime tmSdfExt(llSpotIndex).iTime(0), tmSdfExt(llSpotIndex).iTime(1), "M", "1", slTime         'time of avail in military time
'                        tmSpotDetail(llUpper).sSpotTimeSch = Trim$(Mid$(slTime, 1, 2)) & ":" & Trim$(Mid$(slTime, 3, 5))
'                        tmSpotDetail(llUpper).sSpotTimeAired = Trim$(Mid$(slTime, 1, 2)) & ":" & Trim$(Mid$(slTime, 3, 5))
'                        tmSpotDetail(llUpper).sSpotStatus = tmSdfExt(llSpotIndex).sSchStatus            'scheduled, missed, mg, outside, cancelled, hidden
'                        'Spot type: A=Contract spot; O=Open BB; C=Close BB; L=Cmml Promo;
'                        'N=Feed; S=PSA; M=Promo; D=Donut; B=Bookend; Y=BB only; E=Any BB;
'                        'T=Remnant (if spfSchdRemnant = "N"); Q=per Inquire
'                        'X=Extra Spot
'                        If tmSdfExt(llSpotIndex).sSpotType = "X" Then
'                            tmSpotDetail(llUpper).sSpotType = "F"                           'fill
'                            tmSpotDetail(llUpper).sSpotStatus = "S"                       'fills will be considered a scheduled spot, not an outside since its never been ordered
'                        ElseIf tmSdfExt(llSpotIndex).sSpotType = "O" Then                   'open
'                            tmSpotDetail(llUpper).sSpotType = "O"
'                        ElseIf tmSdfExt(llSpotIndex).sSpotType = "C" Then                   'close
'                            tmSpotDetail(llUpper).sSpotType = "C"
'                        Else
'                            tmSpotDetail(llUpper).sSpotType = "R"                           'regular spot
'                        End If
'                        tmSpotDetail(llUpper).sGame = Trim$(Str$(tmSdfExt(llSpotIndex).iGameNo))    'Game #
'                        tmSpotDetail(llUpper).sSpotLength = Trim$(Str$(tmSdfExt(llSpotIndex).iLen))    'spot length
'                        tmSpotDetail(llUpper).sCopyFlag = tmClf.sLiveCopy               'Live, Pre-recorded, recorded
'                        llRecPos = tmSdfExt(llSpotIndex).lRecPos
'                        llAmount = 0
'                        'if open, close or fill spot, no rate, no need to get the flight
'                        If tmSdfExt(llSpotIndex).sSpotType <> "X" And tmSdfExt(llSpotIndex).sSpotType <> "O" And tmSdfExt(llSpotIndex).sSpotType <> "C" Then
'                            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                            If ilRet = BTRV_ERR_NONE Then
'                                'get rate from flight
'                                ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tlCff)
'                                If ilRet Then
'                                    llAmount = tlCff.lActPrice
'                                End If
'                            End If
'                        End If
'                        tmSpotDetail(llUpper).sGrossSpotPrice = gLongToStrDec(llAmount, 2)
'                        If tgChf.iAgfCode = 0 Then          'direct
'                            ilCommPct = 10000                'no commission
'                        Else
'                            ilCommPct = (10000 - tgCommAgf(ilAgfInx).iCommPct)
'                        End If
'                        slAmount = gLongToStrDec(llAmount, 2)
'                        slSharePct = gIntToStrDec(ilCommPct, 4)                     'net pct (normally 85%)
'                        slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
'                        llNet = gStrDecToLong(slStr, 2)
'                        tmSpotDetail(llUpper).sNetAmt = gRoundStr(slStr, ".01", 2)              'net
'                        tmSpotDetail(llUpper).sGrossSpotPrice = gLongToStrDec(llAmount, 2)      'gross
'                        llComm = llAmount - llNet
'                        tmSpotDetail(llUpper).sAgyCommAmt = gLongToStrDec(llComm, 2)    'net

                        ilSpotTime(0) = tmSdfExt(llSpotIndex).iTime(0)
                        ilSpotTime(1) = tmSdfExt(llSpotIndex).iTime(1)
                        mGatherSpotDetail llUpper, llSpotIndex, ilAgfInx, ilReadSdf, ilSpotTime()
                    
                        'if the spot was an outside or makegood, make sure the schedule line is reread for the next spot since it may have been scheduled in a different vehicle
                        'than it was originally ordered on
                        If tmSdfExt(llSpotIndex).sSchStatus = "G" Or tmSdfExt(llSpotIndex).sSchStatus = "O" Then
                            tmClf.iLine = -1
                            ilPrevVefCode = 0
                        End If
                        llUpper = llUpper + 1
                        ReDim Preserve tmSpotDetail(0 To llUpper) As SPOTDETAIL
 
                    Else                            'in the past; build array of the vehicles in the past to process
                        llCount = llCount + 1
                        ilFound = False
                        For ilLoopOnPast = LBound(imVehiclesInPast) To UBound(imVehiclesInPast) - 1
                            If imVehiclesInPast(ilLoopOnPast) = tmSdfExt(llSpotIndex).iVefCode Then
                                ilFound = True
                                Exit For
                    End If
                        Next ilLoopOnPast
                        If Not ilFound Then
                            imVehiclesInPast(UBound(imVehiclesInPast)) = tmSdfExt(llSpotIndex).iVefCode
                            ReDim Preserve imVehiclesInPast(0 To UBound(imVehiclesInPast) + 1) As Integer
                        End If
                    End If
                  Next llLoopOnSpots
                  mWriteExportSpots         'process last spot
              End If
            End If
        Next ilCurrentRecd
        
        ilAnyPastSpotErrors = mCreatePastSpots(llTodayDate, ilAllDatesInPast)
        
        Erase tlChfAdvtExt
        Erase tmSdfExtSort
        Erase tmSdfExt
        Erase imVehiclesInPast
        
        mCreateSpots = False
        If ilAnySpotErrors Or ilAnyPastSpotErrors Then
            mCreateSpots = True
        End If

End Function

'                   mWriteExportSpots - create a record for every spot from requested dates
'                   Create file Spot-mmddyy-mmddyy.txt (pipe character delimited)
'
'Field   Type    Max Format and example
'1.  Vehicle Alpha   40  Raw Dog
'2.  Vehicle ID  Number  10  Internal CSI #
'3.  Spot Date   Alpha-numeric   8   mm/dd/yy, 1/1/11
'4.  Spot Day    A   3   MON, TUE, WED, THU, FRI, SAT, or SUN
'5.  Spot Time - Scheduled   A-N 8   hh:mm:ss  military time
'6.  Spot Time - Aired   A-N 8   hh:mm:ss  military time (for future spots, aired=scheduled)
'7.  Spot Status A   1   S=scheduled spot, M=missed, G=makegood,  O=Outside, C=cancelled spot, H=Hidden spot
'8.  Spot Type   A   1   R=Regular spot, O=Open Billboard, C=Close billboard, F=Fill
'9.  Contract Type   A   1   C=Regular, V=Reservation, T=Remnant, R=Direct Response, Q=Per Inquiry
'10. Game Number (for Sports)    N   3   001
'11. Spot Length N   3   10, 30, 60, 90, 120, etc.
'12. Advertiser Name A-N 30  Ford
'13. Product Name    A-N 35  Mustang
'14. Agency Name A-N 40  Batton, Barton, Durstin & Osbourne, or blank if a Direct Buy
'15. Contract Number N   9
'16. Primary Competitive Name    A-N 20  Automotive
'17. Secondary Competitive Name  A-N 20  Auto Parts
'18. Gross Spot Price    A-N 11  Show in pennies, including period, but no commas
'19. Agency Commission Amount    A-N 11  Show in pennies, including period, but no commas
'20. Net Amount  A-N 11  Show in pennies, including period, but no commas
'21. Salesperson 1 First Name    A   20
'22. Salesperson 1 Last Name A   20
'23. Salesperson 1 Split %   A-N 8   nnn.nnnn, 4 decimal places, including period, 30.0000 = 30%
'24. Sales Office    A-N 20
'25. Percent ofTrade N   3   0=all cash, 100=all trade
'26. Live/Prerecorded/Recorded   A   1   L or P or R
'27. Group code 1 (Subtotals)    A-N 20  Special Descriptors, for any use
'28. Group code 2 (Markets)  A-N 20  Special Descriptors, for any use
'29. Group code 3 (Formats)  A-N 20  Special Descriptors, for any use
'30. Group code 4 (Research) A-N 20  Special Descriptors, for any use
'31. Group code 5 (Subcompanies) A-N 20  Special Descriptors, for any use

Private Function mWriteExportSpots() As Integer

    Dim ilLoop As Integer
    Dim slRecord As String
    Dim ilIndex As Integer
    Dim slVehicle As String
    Dim ilError As Integer
    Dim ilRecdLen As Integer
    Dim slMonth As String
    Dim slYear As String
    Dim slStr As String
    Dim llAvgRate As Long
    Dim ilRet As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim slAvailTime As String
    Dim slAvailName As String
    Dim slDate As String
    
        DoEvents
        ilError = False
       
        For ilLoop = LBound(tmSpotDetail) To UBound(tmSpotDetail) - 1
            'format the month info for a contract/vehicle
            slRecord = Trim$(tmSpotVefInfo.sVehicleName) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotVefInfo.sVehicleCode) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sSpotDate) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sSpotDay) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sSpotTimeSch) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sSpotTimeAired) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sSpotStatus) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sSpotType) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sContractType) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sGame) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sSpotLength) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sAdvtName) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sProduct) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sAgyName) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sContract) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sPrimComp) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sSecComp) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sGrossSpotPrice) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sAgyCommAmt) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sNetAmt) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sSlspFirstName) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sSlspLastName) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sSlspSplitPct) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sOffice) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotHeader.sTradePct) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotDetail(ilLoop).sCopyFlag) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotVefInfo.sVGGroup1) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotVefInfo.sVGGroup2) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotVefInfo.sVGGroup3) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotVefInfo.sVGGroup4) & DELIMITER
            slRecord = slRecord & Trim$(tmSpotVefInfo.sVGGroup5) & DELIMITER
       
            ilRecdLen = Len(Trim(slRecord))
            ilRet = 0
            On Error GoTo mWriteExportSpotErr
            Print #hmSpots, Left(slRecord, ilRecdLen)
            On Error GoTo 0
            If ilRet <> 0 Then
                imExporting = False
                gLogMsg "Error writing to Export Spot File, Error # " & Str$(ilRet), "ExportInvSpots.txt", False
                mWriteExportSpots = False
                Exit Function
            End If
        Next ilLoop
    
    
        mWriteExportSpots = True
    Exit Function

mWriteExportSpotErr:
    ilRet = Err.Number
    Resume Next
End Function

Private Sub mInitSpotFields()
        'initalize header information for a spot
        tmSpotHeader.sContractType = ""
        tmSpotHeader.sAdvtName = ""
        tmSpotHeader.sProduct = ""
        tmSpotHeader.sAgyName = ""
        tmSpotHeader.sContract = ""
        tmSpotHeader.sPrimComp = ""
        tmSpotHeader.sSecComp = ""
        tmSpotHeader.sSlspFirstName = ""
        tmSpotHeader.sSlspLastName = ""
        tmSpotHeader.sSlspSplitPct = ""
        tmSpotHeader.sOffice = ""
        tmSpotHeader.sTradePct = ""
        
        'initialize common vehicle information
        tmSpotVefInfo.sVehicleName = ""
        tmSpotVefInfo.sVehicleCode = ""
        tmSpotVefInfo.sVGGroup1 = ""
        tmSpotVefInfo.sVGGroup2 = ""
        tmSpotVefInfo.sVGGroup3 = ""
        tmSpotVefInfo.sVGGroup4 = ""
        tmSpotVefInfo.sVGGroup5 = ""
        
        'initialize detail spot information
        ReDim tmSpotDetail(0 To 0) As SPOTDETAIL
        tmSpotDetail(0).sSpotDate = ""
        tmSpotDetail(0).sSpotDay = ""
        tmSpotDetail(0).sSpotTimeSch = ""
        tmSpotDetail(0).sSpotTimeAired = ""
        tmSpotDetail(0).sSpotStatus = ""
        tmSpotDetail(0).sSpotType = ""
        tmSpotDetail(0).sGame = ""
        tmSpotDetail(0).sSpotLength = ""
        tmSpotDetail(0).sGrossSpotPrice = ""
        tmSpotDetail(0).sAgyCommAmt = ""
        tmSpotDetail(0).sNetAmt = ""
        tmSpotDetail(0).sCopyFlag = ""
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

'               mGetSlspFields - retrieve the slsp name, percent of split and office once per contract
'               <input> primary salesperson code from contract header
'               <Return> place into common array for header information:  primary first/last slsp name, % of split (4 dec positions), office name
Private Sub mGetSlspFields()
Dim ilInx As Integer
Dim ilLoop As Integer

        ilInx = gBinarySearchSlf(tgChf.iSlfCode(0))
        If ilInx >= 0 Then
            tmSpotHeader.sSlspFirstName = Trim$(tgMSlf(ilInx).sFirstName)
            tmSpotHeader.sSlspLastName = Trim$(tgMSlf(ilInx).sLastName)
            tmSpotHeader.sSlspSplitPct = gLongToStrDec(tgChf.lComm(0), 4)
        End If
        
        For ilLoop = LBound(tmSofArray) To UBound(tmSofArray) - 1
            If tgMSlf(ilInx).iSofCode = tmSofArray(ilLoop).iCode Then
                tmSpotHeader.sOffice = tmSofArray(ilLoop).sName
                Exit For
            End If
        Next ilLoop
       
 
        Exit Sub
End Sub
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
'           mFindUnReferSpots - find spots that have not been processed due to not an active contract with period requested
'           <input> tlChfAdvtExt()
'
'           Create an array of the contracts to process, sort the array by contract code.
'           Read all the spots by date to see which spots will not be processed because it wasnt active in the
'           requested period.  For example, spots were found in 2020 but a contract existed that had expired in 2012 or 2011.
'           Those spots should be removed from the database.
'
Private Sub mFindUnReferSpots(tlChfAdvtExt() As CHFADVTEXT)
Dim ilCurrentRecd As Integer
Dim slCode As String
ReDim tmChfFound(0 To UBound(tlChfAdvtExt) - LBound(tlChfAdvtExt))
ReDim tlChfNotFound(0 To 0) As SDF
Dim llUpper As Long
Dim llDate As Long
Dim slStartDate As String
Dim slEndDate As String
Dim slStr As String
Dim ilLoop As Integer
Dim slDate As String
Dim ilVefInx As Integer
Dim slVefName As String
Dim ilRet As Integer
Dim ilIndex As Integer
        ilIndex = 0
        For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1
            tmChfFound(ilIndex).lCode = tlChfAdvtExt(ilCurrentRecd).lCode
            tmChfFound(ilIndex).lCntrNo = tlChfAdvtExt(ilCurrentRecd).lCntrNo
            slStr = Trim$(Str(tlChfAdvtExt(ilCurrentRecd).lCode))
            Do While Len(slStr) < 10
                slStr = "0" & slStr
            Loop
            tmChfFound(ilIndex).sKey = Trim$(slStr)
        Next ilCurrentRecd
        'sort the contract codes for binary search by contract code from the spot
        llUpper = UBound(tmChfFound)
        If llUpper > 0 Then
            ArraySortTyp fnAV(tmChfFound(), 0), llUpper, 0, LenB(tmChfFound(0)), 0, LenB(tmChfFound(0).sKey), 0
        End If

        'loop thru the sdf file by date to see which spots wont be referenced
        For llDate = lmUserStartDate To lmUserEndDate
            slStartDate = Format$(llDate, "m/d/yy")
            slEndDate = Format$(llDate, "m/d/yy")
            mObtainSDFByDate hmSdf, slStartDate, slEndDate, tlChfNotFound()
        Next llDate
        
        'cycle thru contracts that are unreferenced and log an error message
        For ilLoop = LBound(tlChfNotFound) To UBound(tlChfNotFound) - 1
            tmSdf = tlChfNotFound(ilLoop)
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
            slVefName = "Unknown vehicle "
            ilVefInx = gBinarySearchVef(tmSdf.iVefCode)
            If ilVefInx >= 0 Then
                slVefName = tgMVef(ilVefInx).sName
            End If
            tmChfSrchKey.lCode = tmSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            lbcInfo(1).AddItem "Contract #" & Trim$(Str$(tmChf.lCntrNo)) & " found on " & Trim$(slDate) & " in " & Trim$(slVefName) & " :check cnt parameters"
            gLogMsg "Contract #" & Trim$(Str$(tmChf.lCntrNo)) & " found on " & Trim$(slDate) & " in " & Trim$(slVefName) & " :check cnt paramters", "ExportInvSpots.txt", False
          
        Next ilLoop
        Erase tlChfNotFound
        Erase tmChfFound
        Exit Sub
End Sub

Public Function mBinarySearchChfFound(llCode As Long) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tmChfFound)
    ilMax = UBound(tmChfFound) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If llCode = tmChfFound(ilMiddle).lCode Then
            'found the match
            mBinarySearchChfFound = ilMiddle
            Exit Function
        ElseIf llCode < tmChfFound(ilMiddle).lCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchChfFound = -1
End Function

Private Sub mObtainSDFByDate(hlSdf As Integer, slStartDate As String, slEndDate As String, llChfNotFound() As SDF)
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE      '7-19-04
    Dim slStr As String
    Dim ilInclude As Integer
    Dim ilChfFoundInx As Integer
    Dim ilFound As Integer


        btrExtClear hlSdf   'Clear any previous extend operation
        ilExtLen = Len(tmSdf)
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
        btrExtClear hlSdf   'Clear any previous extend operation


          gPackDate slStartDate, tmSdfSrchKey4.iDate(0), tmSdfSrchKey4.iDate(1)
          tmSdfSrchKey4.lChfCode = 0
          ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, Len(tmSdf), tmSdfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    
            If ilRet <> BTRV_ERR_END_OF_FILE Then
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Sub
                End If
                Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
      
                ilOffSet = gFieldOffset("Sdf", "sdfChfCode")
                tlLongTypeBuff.lCode = 0                    'match on selective contract code
                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLongTypeBuff, 4)
                If slStartDate <> "" Then
                    gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Sdf", "SdfDate")
                    If slEndDate <> "" Then
                        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                    Else
                        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                    End If
                End If
                If slEndDate <> "" Then
                    gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Sdf", "SdfDate")
                    ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                End If
                ilRet = btrExtAddField(hlSdf, 0, ilExtLen)  'Extract Name
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Sub
                End If
                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                        Exit Sub
                    End If
                    'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                    Loop
                    Do While ilRet = BTRV_ERR_NONE
                        ilChfFoundInx = mBinarySearchChfFound(tmSdf.lChfCode)
                        If ilChfFoundInx = -1 Then          'contract not found
                            'see if already in list
                            ilFound = False
                            For ilChfFoundInx = LBound(llChfNotFound) To UBound(llChfNotFound) - 1
                                
                                If tmSdf.lChfCode = llChfNotFound(ilChfFoundInx).lChfCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilChfFoundInx
                            If Not (ilFound) Then
                                llChfNotFound(UBound(llChfNotFound)) = tmSdf
                                ReDim Preserve llChfNotFound(LBound(llChfNotFound) To UBound(llChfNotFound) + 1) As SDF
                            End If
                        End If
                        ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                        Loop
                    Loop
                End If
            End If              'ilRet <> BTRV_ERR_END_OF_FILE

    Exit Sub
End Sub
'
'               process the spots from earliest date to today (use Today because of posting live log)
'               Those days have potentially been posted
'               mCreatePastSpots
'               <input>  Todays date
'                        ilAllDatesInPast - true if user changing dates to all in the past
'               Return - true if any errors found
Public Function mCreatePastSpots(llTodaysDate As Long, ilAllDatesInPast As Integer) As Integer
    Dim ilRet As Integer
    Dim ilAnyPastSpotErrors As Integer
    Dim ilLoopOnEvt As Integer
    Dim llUpper As Long                 'index to array of sdf records (tmSdfext, always only has 1 for the past processing of spots)
    Dim llEarliestLcfDate As Long
    Dim llLatestLcfDate As Long
    Dim llDate As Long
    Dim ilDate(0 To 1) As Integer
    Dim llTime As Long
    Dim slDate As String
    Dim ilType As Integer
    Dim ilLoopOnVehicles As Integer
    Dim llAvailTime As Long
    Dim ilAvailTime(0 To 1) As Integer
    Dim ilSpot As Integer
    Dim ilInx As Integer
    Dim ilAdfInx As Integer
    Dim ilAgfInx As Integer
    Dim ilReadSdf As Boolean
    Dim llSpotIndex As Long     'index to array of spot info (tmspotdetail) to be written to file
    Dim ilClf As Integer
    Dim llContrCode As Long
    Dim slAvailTime As String
    Dim llAlteredEndDate As Long

    ilAnyPastSpotErrors = 0             'assume everything is OK
    ilReadSdf = False
    'process a vehicle at a time for the spots found that were in the past
    'do all dates starting from start date requested to todays date, unless user has changed
    'dates and they are all in the past
    llAlteredEndDate = llTodaysDate
    If ilAllDatesInPast Then
        llAlteredEndDate = lmUserEndDate
    End If
    For ilLoopOnVehicles = LBound(imVehiclesInPast) To UBound(imVehiclesInPast) - 1
        ilInx = gBinarySearchVef(imVehiclesInPast(ilLoopOnVehicles))
        If ilInx <> -1 Then
            tmVef = tgMVef(ilInx)
            ReDim tmSdfExt(0 To 1) As SDFEXT
            'show Processing vehicle & dates on caption screen & file for the past (sent to routine is already today date minus -1)
            gLogMsg "Processing spots in the past for " & Trim$(tmVef.sName) & " " & Format(lmUserStartDate, "m/d/yy") & " - " & Format(llAlteredEndDate, "m/d/yy"), "ExportInvSpots.txt", False
            lbcInfo(1).AddItem "Processing Spots in the past for " & Trim$(tmVef.sName) & " " & Format(lmUserStartDate, "m/d/yy") & " - " & Format(llAlteredEndDate, "m/d/yy")
            lbcInfo(1).Visible = True
            ilType = 0
            'use the ssf for inventory only
            'For llDate = lmUserStartDate To llTodaysDate
            For llDate = lmUserStartDate To llAlteredEndDate
                DoEvents
                
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
                    ilLoopOnEvt = 1
                    Do While ilLoopOnEvt <= tmSsf.iCount
                        DoEvents
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilLoopOnEvt)
                        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                            'Save time of the current avail
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilLoopOnEvt)
                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAvailTime
                            ilAvailTime(0) = tmAvail.iTime(0)
                            ilAvailTime(1) = tmAvail.iTime(1)
                        
                            gUnpackTime ilAvailTime(0), ilAvailTime(1), "M", "1", slAvailTime         'time of avail in military time

                            For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                               LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilLoopOnEvt + ilSpot)
                                
                                ReDim tmSpotDetail(0 To 0) As SPOTDETAIL
                                llUpper = 0
                                tmSdfSrchKey3.lCode = tmSpot.lSdfCode   'get the spot record
                                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                If (ilRet = BTRV_ERR_NONE) And (tmSdf.sSpotType <> "S" And tmSdf.sSpotType <> "M") And (lmCntrCode = tmSdf.lChfCode Or lmCntrCode = 0) Then  'ignore psa(S) and promo(M)
                                
                                    mInitSpotFields

                                    llContrCode = tmSdf.lChfCode
                                    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())       'read in entire contract
                            
                                    If Not ilRet Then
                                        gLogMsg "** Unable to obtain Contract Code " & Trim$(Str$(llContrCode)), "ExportInvSpots.txt", False
                                        ilAnyPastSpotErrors = True
                                    Else
                                        ReDim tmSpotDetail(0 To 0) As SPOTDETAIL
                                        'put together the header information.  Gather once per contract
                                        llUpper = 0
                                        tmSpotHeader.sContract = Trim$(Str$(tgChf.lCntrNo))
                                        tmSpotHeader.sContractType = tgChf.sType        'regular, reservation, remnant, DR, PI
                                        ilAdfInx = gBinarySearchAdf(tgChf.iAdfCode)
                                        If ilAdfInx >= 0 Then
                                            tmSpotHeader.sAdvtName = Trim$(tgCommAdf(ilAdfInx).sName)
                                        End If
                                        tmSpotHeader.sProduct = tgChf.sProduct
                                        If tgChf.iAgfCode > 0 Then                  'no internal code indicates direct advt
                                            ilAgfInx = gBinarySearchAgf(tgChf.iAgfCode)
                                            If ilAgfInx >= 0 Then
                                                tmSpotHeader.sAgyName = Trim$(tgCommAgf(ilAgfInx).sName)
                                            End If
                                        End If
                                        tmSpotHeader.sPrimComp = mSearchCompCode(tgChf.iMnfComp(0))         'primary competitive code
                                        tmSpotHeader.sSecComp = mSearchCompCode(tgChf.iMnfComp(1))          'secondary competitive code
                                        tmSpotHeader.sTradePct = Trim$(Str$(tgChf.iPctTrade))
                                        mGetSlspFields
                                        tmSpotVefInfo.sVehicleName = Trim$(tmVef.sName)
                                        tmSpotVefInfo.sVehicleCode = Trim$(Str(tmVef.iCode))        'vehicle internal code
                                        tmSpotVefInfo.sVGGroup1 = mGetVGName(tmVef.iMnfVehGp2)            'subtotals
                                        tmSpotVefInfo.sVGGroup2 = mGetVGName(tmVef.iMnfVehGp3Mkt)            'markets
                                        tmSpotVefInfo.sVGGroup3 = mGetVGName(tmVef.iMnfVehGp4Fmt)            'formats
                                        tmSpotVefInfo.sVGGroup4 = mGetVGName(tmVef.iMnfVehGp5Rsch)            'research
                                        tmSpotVefInfo.sVGGroup5 = mGetVGName(tmVef.iMnfVehGp6Sub)            'subcompanies
                                            
                                        'setup common buffer for the spot
                                        tmSdfExt(0).iAdfCode = tgChf.iAdfCode
                                        tmSdfExt(0).iDate(0) = tmSdf.iDate(0)
                                        tmSdfExt(0).iDate(1) = tmSdf.iDate(1)
                                        tmSdfExt(0).iGameNo = tmSdf.iGameNo
                                        tmSdfExt(0).iLen = tmSdf.iLen
                                        tmSdfExt(0).iLineNo = tmSdf.iLineNo
                                        tmSdfExt(0).iTime(0) = tmSdf.iTime(0)
                                        tmSdfExt(0).iTime(1) = tmSdf.iTime(1)
                                        tmSdfExt(0).iVefCode = tmSdf.iVefCode
                                        tmSdfExt(0).lChfCode = tmSdf.lChfCode
                                        tmSdfExt(0).lCode = tmSdf.lCode
                                        tmSdfExt(0).lMdDate = 0             'n/a
                                        tmSdfExt(0).iStatus = 0             'n/a
                                        tmSdfExt(0).sPriceType = tmSdf.sPriceType
                                        tmSdfExt(0).sSchStatus = tmSdf.sSchStatus
                                        tmSdfExt(0).sSpotType = tmSdf.sSpotType
                                        ReDim tmSpotDetail(0 To 1) As SPOTDETAIL
                                        llUpper = 0
                                        llSpotIndex = 0
                                        
                                        For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                                            tmClf = tgClf(ilClf).ClfRec
                                            If tmSdfExt(llSpotIndex).iLineNo = tmClf.iLine Then
                                                mGatherSpotDetail llUpper, llSpotIndex, ilAgfInx, ilReadSdf, ilAvailTime()
                                                mWriteExportSpots         'process last spot

                                                Exit For
                                            End If
                                        Next ilClf

                                    End If
                                 End If                                          'test cntr #
                                'ilLoopOnEvt = ilLoopOnEvt + 1                   'next spot
                            Next ilSpot
                            ilLoopOnEvt = ilLoopOnEvt + tmAvail.iNoSpotsThis                   'bypass spots
                        End If
                        ilLoopOnEvt = ilLoopOnEvt + 1
                    Loop
                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If tmVef.sType = "G" Then
                        ilType = tmSsf.iType
                    End If
                Loop
                slDate = Format$(llDate, "m/d/yy")
                
            Next llDate
        Else                    'didnt find vehicle options
            ilAnyPastSpotErrors = 1
        End If
    Next ilLoopOnVehicles
    
    mCreatePastSpots = ilAnyPastSpotErrors
End Function
'
'                   mGatherSpotDetail - get the information from spot to write the Detail spot records
'                   <input> llUpper (tmSpotDetail)- Detail spot index into list of spots (for the past, theres only 1 record in the array, as each spot is processed one at a time)
'                           llSpotIndex - array of spots from a contract or just one from the avail (tmsdfext)
'                           ilAgfInx = index into the agency record
'                           ilReadSDF - ok to read sdf, SDF already in memory
Public Sub mGatherSpotDetail(llUpper As Long, llSpotIndex As Long, ilAgfInx As Integer, ilReadSdf As Boolean, ilAvailTime() As Integer)
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
                        
                        
                slDayOfWeek = "MONTUEWEDTHUFRISATSUN"
         
                gUnpackDate tmSdfExt(llSpotIndex).iDate(0), tmSdfExt(llSpotIndex).iDate(1), slDate
                tmSpotDetail(llUpper).sSpotDate = slDate  'spot date
                ilDay = gWeekDayStr(slDate)
                tmSpotDetail(llUpper).sSpotDay = Mid$(slDayOfWeek, (ilDay * 3) + 1, 3)      'spot date
                                
                'original scheduled time
                gUnpackTime ilAvailTime(0), ilAvailTime(1), "M", "1", slTime         'time of avail in military time
                tmSpotDetail(llUpper).sSpotTimeSch = Trim$(Mid$(slTime, 1, 2)) & ":" & Trim$(Mid$(slTime, 3, 5))
                
                gUnpackTime tmSdfExt(llSpotIndex).iTime(0), tmSdfExt(llSpotIndex).iTime(1), "M", "1", slTime         'aired time
                tmSpotDetail(llUpper).sSpotTimeAired = Trim$(Mid$(slTime, 1, 2)) & ":" & Trim$(Mid$(slTime, 3, 5))
                tmSpotDetail(llUpper).sSpotStatus = tmSdfExt(llSpotIndex).sSchStatus            'scheduled, missed, mg, outside, cancelled, hidden
                'Spot type: A=Contract spot; O=Open BB; C=Close BB; L=Cmml Promo;
                'N=Feed; S=PSA; M=Promo; D=Donut; B=Bookend; Y=BB only; E=Any BB;
                'T=Remnant (if spfSchdRemnant = "N"); Q=per Inquire
                'X=Extra Spot
                If tmSdfExt(llSpotIndex).sSpotType = "X" Then
                    tmSpotDetail(llUpper).sSpotType = "F"                           'fill
                    tmSpotDetail(llUpper).sSpotStatus = "S"                       'fills will be considered a scheduled spot, not an outside since its never been ordered
                ElseIf tmSdfExt(llSpotIndex).sSpotType = "O" Then                   'open
                    tmSpotDetail(llUpper).sSpotType = "O"
                ElseIf tmSdfExt(llSpotIndex).sSpotType = "C" Then                   'close
                    tmSpotDetail(llUpper).sSpotType = "C"
                Else
                    tmSpotDetail(llUpper).sSpotType = "R"                           'regular spot
                End If
                tmSpotDetail(llUpper).sGame = Trim$(Str$(tmSdfExt(llSpotIndex).iGameNo))    'Game #
                tmSpotDetail(llUpper).sSpotLength = Trim$(Str$(tmSdfExt(llSpotIndex).iLen))    'spot length
                tmSpotDetail(llUpper).sCopyFlag = tmClf.sLiveCopy               'Live, Pre-recorded, recorded
           
                llRecPos = tmSdfExt(llSpotIndex).lRecPos
                llAmount = 0
                'if open, close or fill spot, no rate, no need to get the flight
                If tmSdfExt(llSpotIndex).sSpotType <> "X" And tmSdfExt(llSpotIndex).sSpotType <> "O" And tmSdfExt(llSpotIndex).sSpotType <> "C" Then
                    ilRet = BTRV_ERR_NONE
                    If ilReadSdf Then
                        ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    End If
                    If ilRet = BTRV_ERR_NONE Then
                        'get rate from flight
                        ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tlCff)
                        If ilRet Then
                            llAmount = tlCff.lActPrice
                        End If
                    End If
                End If
                tmSpotDetail(llUpper).sGrossSpotPrice = gLongToStrDec(llAmount, 2)
                If tgChf.iAgfCode = 0 Then          'direct
                    ilCommPct = 10000                'no commission
                Else
                    ilCommPct = (10000 - tgCommAgf(ilAgfInx).iCommPct)
                End If
                slAmount = gLongToStrDec(llAmount, 2)
                slSharePct = gIntToStrDec(ilCommPct, 4)                     'net pct (normally 85%)
                slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
                llNet = gStrDecToLong(slStr, 2)
                tmSpotDetail(llUpper).sNetAmt = gRoundStr(slStr, ".01", 2)              'net
                tmSpotDetail(llUpper).sGrossSpotPrice = gLongToStrDec(llAmount, 2)      'gross
                llComm = llAmount - llNet
                tmSpotDetail(llUpper).sAgyCommAmt = gLongToStrDec(llComm, 2)    'net
                Exit Sub
End Sub
