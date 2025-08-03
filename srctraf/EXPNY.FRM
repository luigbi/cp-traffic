VERSION 5.00
Begin VB.Form ExpNY 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   2625
   ClientWidth     =   9210
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
   ScaleHeight     =   4395
   ScaleWidth      =   9210
   Begin VB.ListBox lbcAdvt 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   5175
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.CommandButton cmcBlackout 
      Appearance      =   0  'Flat
      Caption         =   "&Blackout"
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
      Left            =   5610
      TabIndex        =   9
      Top             =   3990
      Width           =   1050
   End
   Begin VB.ListBox lbcMsg 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   3675
      MultiSelect     =   2  'Extended
      TabIndex        =   21
      Top             =   765
      Width           =   5385
   End
   Begin VB.CommandButton cmcStartDate 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2085
      Picture         =   "EXPNY.frx":0000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   390
      Width           =   195
   End
   Begin VB.TextBox edcStartDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   1155
      MaxLength       =   10
      TabIndex        =   2
      Top             =   390
      Width           =   930
   End
   Begin VB.PictureBox plcCalendar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   1155
      ScaleHeight     =   1740
      ScaleWidth      =   1965
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton cmcCalUp 
         Appearance      =   0  'Flat
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.CommandButton cmcCalDn 
         Appearance      =   0  'Flat
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   45
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         Width           =   285
      End
      Begin VB.PictureBox pbcCalendar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1440
         Left            =   45
         Picture         =   "EXPNY.frx":00FA
         ScaleHeight     =   1410
         ScaleWidth      =   1845
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   255
         Width           =   1875
         Begin VB.Label lacDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   510
            TabIndex        =   18
            Top             =   390
            Visible         =   0   'False
            Width           =   300
         End
      End
      Begin VB.Label lacCalName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   330
         TabIndex        =   5
         Top             =   45
         Width           =   1305
      End
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   15
      Width           =   3435
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3135
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2370
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2790
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   150
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   765
      Width           =   3375
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
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
      Left            =   2625
      TabIndex        =   7
      Top             =   3990
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
      Left            =   4095
      TabIndex        =   8
      Top             =   3990
      Width           =   1050
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   3870
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacProcessing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   135
      TabIndex        =   20
      Top             =   3615
      Width           =   8895
   End
   Begin VB.Label lacStartDate 
      Appearance      =   0  'Flat
      Caption         =   "Export Date"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   1065
   End
   Begin VB.Label lacErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5055
      TabIndex        =   15
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   5175
      TabIndex        =   13
      Top             =   1395
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lacCntr 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3225
      TabIndex        =   11
      Top             =   1395
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ExpNY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of EXPNY.FRM on Fri 3/12/10 @ 11:00 AM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ExpNY.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Export feed (for New York) input screen code
Option Explicit
Option Compare Text
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmTo As Integer   'From file hanle
Dim hmMsg As Integer   'From file hanle
Dim lmMaxWidth As Long
Dim lmNowDate As Long   'Todays date
'Required by gMakeSsf
Dim tmSsf As SSF                'SSF record image
Dim hmSsf As Integer
Dim hmCTSsf As Integer
Dim tmCTSsf As SSF               'Ssf for conflict test
'Dim tmSsfOld As SSF
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmAvailTest As AVAILSS
Dim tmPrevAvail As AVAILSS
'Advertiser name
Dim hmAdf As Integer
Dim tmAdf As ADF
Dim tmAdfSrchKey As INTKEY0 'ANF key record image
Dim imAdfRecLen As Integer  'ANF record length
Dim tmAdvertiser() As SORTCODE
Dim smAdvertiserTag As String
'Agency name
Dim hmAgf As Integer
Dim tmAgf As AGF
Dim imAgfRecLen As Integer  'AGF record length
'Multi- name
Dim hmMnf As Integer
Dim tmMnf As MNF
Dim imMnfRecLen As Integer  'MNF record length
'Copy/Product
Dim hmCpf As Integer
Dim tmCpf As CPF
Dim tmCpfSrchKey As LONGKEY0 'CPEF key record image
Dim imCpfRecLen As Integer  'CPF record length
'Avail name
Dim hmAnf As Integer
Dim tmAnf As ANF
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length
'Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
'Contract Line record information
Dim hmClf As Integer        'Contract Line file handle
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
'Short Title Vehicle Table record information
Dim hmVsf As Integer        'Short Title Vehicle Table file handle
Dim imVsfRecLen As Integer  'VSF record length
Dim tmVsf As VSF            'VSF record image
'Prduct code file
Dim hmPrf As Integer 'Product file handle
Dim tmPrf As PRF        'PRF record image
Dim imPrfRecLen As Integer
'Short Title record information
Dim hmSif As Integer        'Short Title file handle
Dim imSifRecLen As Integer  'SIF record length
Dim tmSif As SIF            'SIF record image
'Media code record information
Dim hmMcf As Integer        'Contract line file handle
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer  'MCF record length
Dim tmMcf As MCF            'MCF record image
'Copy inventory record information
Dim hmCif As Integer        'Copy line file handle
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer  'CIF record length
Dim tmCif As CIF            'CIF record image
Dim hmTzf As Integer        'Time zone Copy file handle
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfRecLen As Integer  'TZF record length
Dim tmTzf As TZF            'TZF record image
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmSVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
Dim smVehName As String
Dim tmUserVehicle() As SORTCODE
Dim smUserVehicleTag As String
'Vehicle linkage record information
Dim hmVlf As Integer        'Vehicle linkage file handle
Dim tmVlfSrchKey1 As VLFKEY1 'VLF key record image
Dim imVlfRecLen As Integer  'VLF record length
Dim tmVlf As VLF            'VLF record image
'Delivery file (DLF)
Dim hmEgf As Integer        'Delivery link file
Dim imEgfRecLen As Integer  'DLF record length
Dim tmEgfSrchKey As DLFKEY0 'DLF key record image
Dim tmEgf As DLF            'DLF record image
'Prefeed
Dim hmPff As Integer
Dim tmPff As PFF        'GSF record image
Dim imPffRecLen As Integer        'GSF record length
Dim tmPffSrchKey0 As LONGKEY0
Dim tmPffSrchKey1 As PFFKEY1
Dim tmPrefeedExpt() As PREFEEDEXPT
Dim hmPbf As Integer
Dim tmPbf As PBF        'GSF record image
Dim imPbfRecLen As Integer        'GSF record length
Dim tmPbfSrchKey0 As LONGKEY0
Dim tmPbfSrchKey1 As LONGKEY0
Dim tmPBFInfo() As PBFINFO
'Spot record
Dim tmSdf As SDF
Dim hmSdf As Integer
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey3 As LONGKEY0
'Blackout record information
Dim hmBof As Integer        'Blackout file handle
Dim imBofRecLen As Integer  'BOF record length
Dim tmBof As BOF            'BOF record image

Dim hmGsf As Integer
Dim tmGsf As GSF        'GSF record image
Dim tmGsfSrchKey3 As GSFKEY3    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

'8/29/14: Check for Airing vehicle copy (ttp 7064)
Dim hmCrf As Integer
Dim hmRsf As Integer
Dim hmCvf As Integer

Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smEASOrASPOrESPN As String
'Calendar
Dim tmCDCtrls(0 To 7) As FIELDAREA
Dim imLBCDCtrls As Integer
Dim imCalYear As Integer    'Month of displayed calendar
Dim imCalMonth As Integer   'Year of displayed calendar
Dim lmCalStartDate As Long  'Start date of displayed calendar
Dim lmCalEndDate As Long    'End date of displayed calendar
Dim imCalType As Integer
Dim imEvtType(0 To 14) As Integer
Dim imCopyMissing As Integer
Dim tmSpotTimes() As SPOTTIMESNY
Dim tmSvSpotTimes() As SPOTTIMESNY
Dim smNewLines() As String * 72
Dim imFirstActivate As Integer
Dim smEngrHideMediaCode As String

' MsgBox parameters
Const vbOkOnly = 0                 ' OK button only
Const vbCritical = 16          ' Critical message
Const vbApplicationModal = 0
Const INDEXKEY0 = 0
Private Sub cmcBlackout_Click()
    'igShowHelpMsg = imShowHelpMsg
    'Blackout.Show vbModal
    Dim slStr As String
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    'Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
        If igTestSystem Then
            slStr = "ExpNY^Test\" & sgUserName
        Else
            slStr = "ExpNY^Prod\" & sgUserName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "ExpNY^Test^NOHELP\" & sgUserName
    '    Else
    '        slStr = "ExpNY^Prod^NOHELP\" & sgUserName
    '    End If
    'End If
    slStr = slStr & "\NY"
    'lgShellRet = Shell(sgExePath & "Blackout.Exe " & slStr, 1)
'    ExpNY.Enabled = False
'    Do While Not igChildDone
'        DoEvents
'    Loop
'    slStr = sgDoneMsg
'    ExpNY.Enabled = True
'    edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
'    For ilLoop = 0 To 10
'        DoEvents
'    Next ilLoop
'    'Screen.MousePointer = vbDefault    'Default
    sgCommandStr = slStr
    Blackout.Show vbModal
End Sub
Private Sub cmcBlackout_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcCalDn_Click()
    imCalMonth = imCalMonth - 1
    If imCalMonth <= 0 Then
        imCalMonth = 12
        imCalYear = imCalYear - 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcCalUp_Click()
    imCalMonth = imCalMonth + 1
    If imCalMonth > 12 Then
        imCalMonth = 1
        imCalYear = imCalYear + 1
    End If
    pbcCalendar_Paint
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    plcCalendar.Visible = False
End Sub
Private Sub cmcExport_Click()
    Dim slToFile As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStr As String
    Dim ilDays As Integer
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slLetter As String
    Dim slFileName As String
    Dim ilLine As Integer
    Dim slDateTime As String
    Dim slMissingCopyNames As String
    If imExporting Then
        Exit Sub
    End If
    On Error GoTo ExportError
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        Beep
        edcStartDate.SetFocus
        Exit Sub
    End If
    slStartDate = slStr
    ilDays = 1  'Val(edcNoDays.Text)
    'If ilDays = 0 Then
    '    Beep
    '    edcNoDays.SetFocus
    '    Exit Sub
    'End If
    slEndDate = Format$(gDateValue(slStartDate) + ilDays - 1, "m/d/yy")
    lbcMsg.Clear
    lmMaxWidth = 0
    If Not mOpenMsgFile() Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    'Determine file name from date: YYMMDDn.NY
    'where n= version letter
    gObtainYearMonthDayStr slStartDate, True, slFYear, slFMonth, slFDay
    slFileName = right$(slFYear, 2) & slFMonth & slFDay
    If smEASOrASPOrESPN = "ASP" Then
        slLetter = "AE"
    ElseIf smEASOrASPOrESPN = "ESPN" Then
        slLetter = "ESPN"
    Else
        slLetter = "CS"
    End If
    ilRet = 0
    'On Error GoTo cmcExportErr:
    slToFile = slFileName & slLetter & ".TXT"
    'slDateTime = FileDateTime(sgExportPath & slToFile)
    ilRet = gFileExist(sgExportPath & slToFile)
    If ilRet = 0 Then
        Kill sgExportPath & slToFile
    End If
    'slLetter = "A"
    'Do
    '    ilRet = 0
    '    On Error GoTo cmcExportErr:
    '    slToFile = slFileName & slLetter & ".NY"
    '    slDateTime = FileDateTime(sgExportPath & slToFile)
    '    If ilRet = 0 Then
    '        slLetter = Chr(Asc(slLetter) + 1)
    '    End If
    'Loop While ilRet = 0
    'slNum = "1"
    'Do
    '    ilRet = 0
    '    On Error GoTo cmcExportErr:
    '    slToFile = slFileName & slNum & ".NY"
    '    slDateTime = FileDateTime(sgExportPath & slToFile)
    '    If ilRet = 0 Then
    '        slNum = Trim$(Str$(Val(slNum) + 1))
    '    End If
    'Loop While ilRet = 0
    On Error GoTo 0
    slToFile = sgExportPath & slFileName & slLetter & ".TXT" 'slToFile
    'ReDim smCurrentLines(0 To 0) As String * 52
    'Obtain file that was last sent
    'On Error GoTo cmcExportErr:
    'If slLetter <> "A" Then
    '    ilRet = 0
    '    hlRead = FreeFile
    '    Open sgExportPath & slFileName & Chr(Asc(slLetter) - 1) & ".NY" For Input Access Read As hlRead
    '    If ilRet = 0 Then
    '        Do While Not EOF(hlRead)
    '            Line Input #hlRead, slLine
    '            smCurrentLines(UBound(smCurrentLines)) = slLine
    '            ReDim Preserve smCurrentLines(0 To UBound(smCurrentLines) + 1) As String * 52
    '        Loop
    '        Close #hlRead
    '    End If
    'End If
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo cmcExportErr:
    'hmTo = FreeFile
    'Open slToFile For Output As hmTo
    ilRet = gFileOpen(slToFile, "Output", hmTo)
    If ilRet <> 0 Then
        'Print #hmMsg, "** Terminated **"
        gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
        Close #hmMsg
        ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
        'edcTo.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ReDim smNewLines(0 To 0) As String * 72
    imExporting = True
    slMissingCopyNames = ""
    ilRet = gReadBofRec(0, hmBof, hmCif, hmPrf, hmSif, hmCHF, "B", slStartDate, 1)
    Screen.MousePointer = vbHourglass
    igStartBofIndex = LBound(tgRBofRec) - 1
    ig30StartBofIndex = igStartBofIndex
    ig60StartBofIndex = igStartBofIndex
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        If lbcVehicle.Selected(ilLoop) Then
            imCopyMissing = False
            slNameCode = tmUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            smVehName = Trim$(slName)
            'Print #hmMsg, "Generating Data for " & Trim$(slName)
            gAutomationAlertAndLogHandler "Generating Data for " & Trim$(slName)
            lacProcessing.Caption = "Generating Data for " & Trim$(slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            If Not mExpNYDay("C", ilVefCode, slStartDate, slEndDate, "12AM", "12AM", imEvtType()) Then
                'Print #hmMsg, "** Terminated **"
                gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
                Close #hmMsg
                Close #hmTo
                imExporting = False
                'MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                Screen.MousePointer = vbDefault
                cmcCancel.SetFocus
                Exit Sub
            End If
            If imCopyMissing Then
                If slMissingCopyNames = "" Then
                    slMissingCopyNames = slName
                Else
                    slMissingCopyNames = slMissingCopyNames & ", " & slName
                End If
            End If
        End If
    Next ilLoop
    lacProcessing.Caption = "Writing Data to " & slFileName & slLetter & ".TXT"
    DoEvents
    ilRet = 0
    'On Error GoTo cmcExportErr:
    For ilLine = LBound(smNewLines) To UBound(smNewLines) - 1 Step 1
        If (smEASOrASPOrESPN = "ASP") Or (smEASOrASPOrESPN = "ESPN") Then
            'Print #hmTo, Trim$(smNewLines(ilLine))
            Print #hmTo, smNewLines(ilLine) 'Output 72 bytes
        Else
            Print #hmTo, Trim$(Left$(smNewLines(ilLine), 52))
        End If
        If ilRet <> 0 Then
            imExporting = False
            'Print #hmMsg, "** Terminated **"
            gAutomationAlertAndLogHandler "** Terminated:" & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            Close #hmTo
            Screen.MousePointer = vbDefault
            ''MsgBox "Error writing to " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            gAutomationAlertAndLogHandler "Error writing to " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
            cmcCancel.SetFocus
            Exit Sub
        End If
    Next ilLine
    'Print #hmMsg, "** Completed **"
    gAutomationAlertAndLogHandler "** Completed **"
    Close #hmMsg
    Close #hmTo
    On Error GoTo 0
    ''Output the difference file
    'lacProcessing.Caption = "Creating Difference Data"
    'DoEvents
    'slToFile = sgExportPath & slFileName & slNum & ".NY" 'slToFile
    'On Error GoTo cmcExportErr:
    'hmTo = FreeFile
    'Open slToFile For Output As hmTo
    'If ilRet <> 0 Then
    '    MsgBox "Open " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
    '    'edcTo.SetFocus
    '    Exit Sub
    'End If
    'If UBound(smCurrentLines) > LBound(smCurrentLines) Then
    '    For ilLoop = LBound(smNewLines) To UBound(smNewLines) - 1 Step 1
    '        slTestLine = Trim$(smNewLines(ilLoop))
    '        slTestLine = Right$(slTestLine, Len(slTestLine) - 1)
    '        For ilLine = LBound(smCurrentLines) + 1 To UBound(smCurrentLines) - 1 Step 1
    '            slCurrentLine = Trim$(smCurrentLines(ilLine))
    '            slCurrentLine = Right$(slCurrentLine, Len(slCurrentLine) - 1)
    '            If StrComp(slTestLine, slCurrentLine, 1) = 0 Then
    '                smCurrentLines(ilLine) = "Removed"
    '                smNewLines(ilLoop) = "Removed"
    '                Exit For
    '            End If
    '        Next ilLine
    '    Next ilLoop
    'End If
    'lacProcessing.Caption = "Writing Data to " & slFileName & slNum & ".NY"
    'DoEvents
    ''Output header record
    'slRecord = "COMM "
    ''Start Date
    'gObtainYearMonthDayStr slStartDate, True, slFYear, slFMonth, slFDay
    'slRecord = slRecord & Right$(slFYear, 2) & slFMonth & slFDay & " "
    ''Start time
    'slRecord = slRecord & "000000" & " "
    ''End Date
    'gObtainYearMonthDayStr slEndDate, True, slFYear, slFMonth, slFDay
    'slRecord = slRecord & Right$(slFYear, 2) & slFMonth & slFDay & " "
    ''Start time
    'slRecord = slRecord & "235959"
    'ilRet = 0
    'On Error GoTo cmcExportErr:
    'Print #hmTo, slRecord
    'On Error GoTo 0
    'If ilRet <> 0 Then
    '    MsgBox "Error writing to " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
    '    cmcCancel.SetFocus
    '    Exit Sub
    'End If
    'For ilLine = LBound(smCurrentLines) + 1 To UBound(smCurrentLines) - 1 Step 1
    '    If Trim$(smCurrentLines(ilLine)) <> "Removed" Then
    '        Mid$(smCurrentLines(ilLine), 1, 1) = "D"
    '        Print #hmTo, smCurrentLines(ilLine)
    '    End If
    'Next ilLine
    'For ilLine = LBound(smNewLines) To UBound(smNewLines) - 1 Step 1
    '    If Trim$(smNewLines(ilLine)) <> "Removed" Then
    '        Print #hmTo, smNewLines(ilLine)
    '    End If
    'Next ilLine
    'Close #hmTo


    lacProcessing.Caption = ""
    Screen.MousePointer = vbDefault
    'If slMissingCopyNames <> "" Then
    '    MsgBox "Copy missing on " & slMissingCopyNames, vbOkOnly + vbExclamation + vbApplicationModal, "Copy Missing"
    'End If
    cmcCancel.Caption = "&Done"
    imExporting = False
    Exit Sub
'cmcExportErr:
'    ilRet = Err.Number
'    Resume Next

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub
Private Sub cmcExport_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub cmcStartDate_Click()
    plcCalendar.Visible = Not plcCalendar.Visible
    edcStartDate.SelStart = 0
    edcStartDate.SelLength = Len(edcStartDate.Text)
    edcStartDate.SetFocus
End Sub
Private Sub cmcStartDate_GotFocus()
    cmcCancel.Caption = "&Cancel"
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub edcStartDate_Change()
    Dim slStr As String
    slStr = edcStartDate.Text
    If Not gValidDate(slStr) Then
        lacDate.Visible = False
        Exit Sub
    End If
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
End Sub
Private Sub edcStartDate_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        'mInitDDE
        'mSendHelpMsg "BT"
    End If
    gCtrlGotFocus edcStartDate
    cmcCancel.Caption = "&Cancel"
End Sub
Private Sub edcStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcStartDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcStartDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYSLASH) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcStartDate_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim slDate As String
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        If (Shift And vbAltMask) > 0 Then
            plcCalendar.Visible = Not plcCalendar.Visible
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYUP Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 7, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 7, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        If (Shift And vbAltMask) > 0 Then
        Else
            slDate = edcStartDate.Text
            If gValidDate(slDate) Then
                If KeyCode = KEYLEFT Then 'Up arrow
                    slDate = Format$(gDateValue(slDate) - 1, "m/d/yy")
                Else
                    slDate = Format$(gDateValue(slDate) + 1, "m/d/yy")
                End If
                gObtainMonthYear imCalType, slDate, imCalMonth, imCalYear
                edcStartDate.Text = slDate
            End If
        End If
        edcStartDate.SelStart = 0
        edcStartDate.SelLength = Len(edcStartDate.Text)
    End If
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    DoEvents    'Process events so pending keys are not sent to this
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_GotFocus()
    plcCalendar.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        plcCalendar.Visible = False
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmAdvertiser
    Erase tmUserVehicle
    Erase tmPBFInfo
    Erase tmSpotTimes
    Erase tmSvSpotTimes
    Erase tmPrefeedExpt
    Erase smNewLines
    Erase tmPrefeedExpt
    Erase tmPBFInfo
    Erase tmAdvertiser
    Erase smNewLines
    Erase tmSpotTimes
    Erase tgSpotSum
    
    '8/29/14: Check for Airing vehicle copy (ttp 7064)
    ilRet = btrClose(hmCvf)
    btrDestroy hmCvf
    ilRet = btrClose(hmRsf)
    btrDestroy hmRsf
    ilRet = btrClose(hmCrf)
    btrDestroy hmCrf
    
    ilRet = btrClose(hmCTSsf)
    btrDestroy hmCTSsf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    ilRet = btrClose(hmSif)
    btrDestroy hmSif
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmTzf)
    btrDestroy hmTzf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    ilRet = btrClose(hmEgf)
    btrDestroy hmEgf
    ilRet = btrClose(hmPff)
    btrDestroy hmPff
    ilRet = btrClose(hmPbf)
    btrDestroy hmPbf
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmBof)
    btrDestroy hmBof
    ilRet = btrClose(hmGsf)
    btrDestroy hmGsf
    
    Set ExpNY = Nothing   'Remove data segment
    
End Sub
'Private Sub imcHelp_Click()
'    Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'    Traffic!cdcSetup.HelpCommand = cdlHelpIndex
'    Traffic!cdcSetup.Action = 6
'End Sub
Private Sub lbcVehicle_GotFocus()
    cmcCancel.Caption = "&Cancel"
    plcCalendar.Visible = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBoxCalDate                     *
'*                                                     *
'*             Created:8/25/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Place box around calendar date *
'*                                                     *
'*******************************************************
Private Sub mBoxCalDate()
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim llInputDate As Long
    Dim ilWkDay As Integer
    Dim slDay As String
    Dim llDate As Long
    slStr = edcStartDate.Text
    If gValidDate(slStr) Then
        llInputDate = gDateValue(slStr)
        If (llInputDate >= lmCalStartDate) And (llInputDate <= lmCalEndDate) Then
            ilRowNo = 0
            llDate = lmCalStartDate
            Do
                ilWkDay = gWeekDayLong(llDate)
                slDay = Trim$(str$(Day(llDate)))
                If llDate = llInputDate Then
                    lacDate.Caption = slDay
                    lacDate.Move tmCDCtrls(ilWkDay + 1).fBoxX - 30, tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) - 30
                    lacDate.Visible = True
                    Exit Sub
                End If
                If ilWkDay = 6 Then
                    ilRowNo = ilRowNo + 1
                End If
                llDate = llDate + 1
            Loop Until llDate > lmCalEndDate
            lacDate.Visible = False
        Else
            lacDate.Visible = False
        End If
    Else
        lacDate.Visible = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mExpNYDay                       *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build export record             *
'*                     ("EST" zone, Cmml Sch = "Y",    *
'*                      and Avail Name = "N")          *
'*                                                     *
'*******************************************************
Private Function mExpNYDay(sLCP As String, ilVefCode As Integer, slSDate As String, slEDate As String, slStartTime As String, slEndTime As String, ilEvtType() As Integer) As Integer
'
'   ilRet = mExpNYDay (slCp, ilVefCode, slSDate, slEDate, slStartTime, slEndTime, ilEvtType())
'
'   Where:
'       slCP (I)- "C"=Current only; "P"=Pending only; "B"=Both
'       ilVefCode (I)-Vehicle code number(slFor = L or C) or feed code (slFor = D)
'       slSDate (I)- Start Date that events are to be obtained
'       slEDate (I)- Start Date that events are to be obtained
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'       ilEvtType (I)- Array of which events are to be included (True or False)
'                       Index description
'                         0   Library
'                         1   Program event
'                         2   Contract avail
'                         3   Open BB
'                         4   Floating BB
'                         5   Close BB
'                         6   Cmml promo
'                         7   Feed avail
'                         8   PSA avail
'                         9   Promo avail
'                         10  Page eject
'                         11  Line space 1
'                         12  Line space 2
'                         13  Line space 3
'                         14  Other event types
'
'
'
    'Dim slType As String
    Dim ilType As Integer
    Dim ilRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilSsfDate0 As Integer
    Dim ilSsfDate1 As Integer
    Dim ilEvt As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim ilSpot As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilTerminated As Integer
    Dim slTime As String
    Dim slLen As String
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim ilVpfIndex As Integer
    Dim ilVehCode As Integer
    Dim ilEgfDate0 As Integer
    Dim ilEgfDate1 As Integer
    Dim ilEgfFound As Integer
    Dim ilVlfDate0 As Integer
    Dim ilVlfDate1 As Integer
    Dim ilAirHour As Integer
    Dim ilFeedHour As Integer
    Dim ilSIndex As Integer
    Dim slSsfDate As String
    'Spot summary
    Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim tlSsfSrchKey2 As SSFKEY2
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim llEvtTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilWithinTime As Integer
    Dim ilEvtFdIndex As Integer
    Dim slAirDate As String
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim llAirDate As Long
    Dim ilTest As Integer
    Dim llSpotTime As Long
    Dim llHour As Long
    Dim llMin As Long
    Dim llSec As Long
    Dim llAvailTime As Long
    Dim llPrevAvailEndTime As Long
    Dim ilEgfExist As Integer
    Dim ilEgfTime0 As Integer
    Dim ilEgfTime1 As Integer
    Dim ilPass As Integer
    Dim ilNoPasses As Integer
    Dim ilPassFound As Integer
    Dim ilCopy As Integer
    'Spot detail record information
    Dim slStr As String
    Dim ilAvailUnits As Integer
    Dim ilAvailLen As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim llCopyMissingSdfCode As Long
    Dim ilSubFeed As Integer    'True=Output only Dlf records with subfeeds;
                                'False=Output only records without subfeeds
    Dim slMsg As String
    Dim slZone As String
    Dim slShortTitle As String * 15
    Dim slCodeStn As String * 2
    Dim slRecord As String
    Dim slISCI As String
    '2/20/15: Moved within Prefeed loop so that running time within breaks restarted
    'ReDim tmSpotTimes(0 To 0) As SPOTTIMESNY
    ReDim tgSpotSum(0 To 0) As SPOTSUM
    Dim slXMid As String
    Dim ilPrefeed As Integer
    Dim blAvailOk As Boolean
    Dim ilAnf As Integer
    Dim blSport As Boolean
    Dim ilVef As Integer
    Dim ilBus As Integer
    Dim slBuses As String
    Dim slBus() As String
    'ReDim smNewLines(0 To 0) As String * 52
    Dim ilVff As Integer
    Dim blBypassZeroUnits As Boolean


    'slType = "O"
    ilType = 0
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    ilSsfRecLen = Len(tmSsf)  'Get and save SSF record length
    On Error GoTo 0
    tmVef.iCode = 0
    llCopyMissingSdfCode = 0
    llSDate = gDateValue(slSDate)
    llEDate = gDateValue(slEDate)
    '4/28/10:  Switch over to use the PreFeed area which includes the actual day
    'For llDate = llSDate - 1 To llEDate Step 1 'Process next date for avails that map back one day
    blSport = False
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType = "G" Then
            blSport = True
        End If
    End If
    If blSport Then
        ReDim tmPrefeedExpt(0 To 1) As PREFEEDEXPT
        gPackDateLong llSDate, ilLogDate0, ilLogDate1
        tmPrefeedExpt(0).iLogDate0 = ilLogDate0
        tmPrefeedExpt(0).iLogDate1 = ilLogDate1
        tmPrefeedExpt(0).iDlfDate0 = 0
        tmPrefeedExpt(0).iDlfDate1 = 0
        tmPrefeedExpt(0).lPffCode = 0
        tmPrefeedExpt(0).lFStartTime = 0
        tmPrefeedExpt(0).lFEndTime = 86400
        tmPrefeedExpt(0).iFDay = -1
        tmPrefeedExpt(0).sFZone = ""
        tmPrefeedExpt(0).lAdjTime = 0
        tmPrefeedExpt(0).iTDay = -1
    Else
        gBuildExportDates hmEgf, hmPff, ilVefCode, "E", llSDate, -1, tmPrefeedExpt()
    End If
    For ilPrefeed = LBound(tmPrefeedExpt) To UBound(tmPrefeedExpt) - 1 Step 1
        '2/20/15: Moved here so that running time within breaks restatrted
        ReDim tmSpotTimes(0 To 0) As SPOTTIMESNY
        mGetPBfRecs tmPrefeedExpt(ilPrefeed).lPffCode
        lgStartIndex = UBound(smNewLines)
        gUnpackDateLong tmPrefeedExpt(ilPrefeed).iLogDate0, tmPrefeedExpt(ilPrefeed).iLogDate1, llDate
        ilWithinTime = False
        slDate = Format$(llDate, "m/d/yy")
        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        ilVehCode = ilVefCode
        If ilVehCode <> tmVef.iCode Then
            tmVefSrchKey.iCode = ilVehCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If (ilRet <> BTRV_ERR_NONE) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox(smVehName & " get error ", vbOkOnly + vbCritical + vbApplicationModal, "Find Error")
                Screen.MousePointer = vbHourglass
                mExpNYDay = False
                Exit Function
            End If
            ilVpfIndex = -1
            'For ilLoop = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
            '    If tmVef.iCode = tgVpf(ilLoop).iVefKCode Then
                ilLoop = gBinarySearchVpf(tmVef.iCode)
                If ilLoop <> -1 Then
                    ilVpfIndex = ilLoop
            '        Exit For
                End If
            'Next ilLoop
            If ilVpfIndex = -1 Then
                mExpNYDay = False
                Screen.MousePointer = vbDefault
                ilRet = MsgBox(smVehName & " preference get error ", vbOkOnly + vbCritical + vbApplicationModal, "Find Error")
                Screen.MousePointer = vbHourglass
                Exit Function
            End If
        End If
        ilEgfFound = False
        'If (tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(1) <> 0)) Then
        If (tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(0) <> 0)) Then
            'Obtain Engineering records for date
            If (ilDay >= 0) And (ilDay <= 4) Then
                slDay = "0"
            ElseIf ilDay = 5 Then
                slDay = "6"
            Else
                slDay = "7"
            End If
            ''Obtain the start date of Egf
            'tmEgfSrchKey.iVefCode = ilVehCode
            'tmEgfSrchKey.sAirDay = slDay
            'tmEgfSrchKey.iStartDate(0) = ilLogDate0
            'tmEgfSrchKey.iStartDate(1) = ilLogDate1
            'tmEgfSrchKey.iAirTime(0) = 0
            'tmEgfSrchKey.iAirTime(1) = 6144 '24*256
            'ilRet = btrGetLessOrEqual(hmEgf, tmEgf, imEgfRecLen, tmEgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            'If (ilRet = BTRV_ERR_NONE) And (tmEgf.iVefCode = ilVehCode) And (tmEgf.sAirDay = slDay) Then
            '    ilEgfDate0 = tmEgf.iStartDate(0)
            '    ilEgfDate1 = tmEgf.iStartDate(1)
            '    ilEgfFound = True
            'Else
            '    ilEgfDate0 = 0
            '    ilEgfDate1 = 0
            'End If
            ilEgfDate0 = tmPrefeedExpt(ilPrefeed).iDlfDate0
            ilEgfDate1 = tmPrefeedExpt(ilPrefeed).iDlfDate1
            If ilEgfDate1 <> 0 Then
                ilEgfFound = True
            Else
                ilEgfFound = False
            End If
            'Obtain the start date of VLF
            If tmVef.sType = "A" Then
                ilVlfDate0 = 0
                ilVlfDate1 = 0
                tmVlfSrchKey1.iAirCode = ilVehCode
                tmVlfSrchKey1.iAirDay = Val(slDay)
                tmVlfSrchKey1.iEffDate(0) = ilLogDate0
                tmVlfSrchKey1.iEffDate(1) = ilLogDate1
                tmVlfSrchKey1.iAirTime(0) = 0
                tmVlfSrchKey1.iAirTime(1) = 6144    '24*256
                tmVlfSrchKey1.iAirPosNo = 32000
                tmVlfSrchKey1.iAirSeq = 32000
                ilRet = btrGetLessOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode)
                    ilTerminated = False
                    'Check for CBS
                    If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                        If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                            ilTerminated = True
                        End If
                    End If
                    If (tmVlf.sStatus <> "P") And (tmVlf.iAirDay = Val(slDay)) And (Not ilTerminated) Then
                        ilVlfDate0 = tmVlf.iEffDate(0)
                        ilVlfDate1 = tmVlf.iEffDate(1)
                        Exit Do
                    End If
                    ilRet = btrGetPrevious(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                'If (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) Then
                '    ilVlfDate0 = tmVlf.iEffDate(0)
                '    ilVlfDate1 = tmVlf.iEffDate(1)
                'Else
                '    ilVlfDate0 = 0
                '    ilVlfDate1 = 0
                'End If
            End If
        End If
        
        blBypassZeroUnits = False
        If (tmVef.sType = "A") Then
            ilVff = gBinarySearchVff(tmVef.iCode)
            If ilVff <> -1 Then
                If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                    blBypassZeroUnits = True
                End If
            End If
        End If
        
        'If Not ilEgfFound Then
        '    MsgBox smVehName & " engineering missing", vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
        '    mExpNYDay = False
        '    Exit Function
        'End If
        'gObtainVlf hlVlf, ilVehCode, llDate, tlVlf0(), tlVlf5(), tlVlf6()
        'If ilEgfFound Then
        If (ilEgfFound) Or (blSport) Then
            llPrevAvailEndTime = -1
            ilDay = gWeekDayStr(slDate)
            gPackDate slDate, ilLogDate0, ilLogDate1
            ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilSsfDate0 = ilLogDate0
            ilSsfDate1 = ilLogDate1
            ''tlSsfSrchKey.sType = slType
            'tlSsfSrchKey.iType = ilType
            'tlSsfSrchKey.iVefCode = ilVehCode
            'tlSsfSrchKey.iDate(0) = ilSsfDate0
            'tlSsfSrchKey.iDate(1) = ilSsfDate1
            'tlSsfSrchKey.iStartTime(0) = 0
            'tlSsfSrchKey.iStartTime(1) = 0
            'ilRet = gSSFGetEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            tlSsfSrchKey2.iVefCode = ilVefCode
            tlSsfSrchKey2.iDate(0) = ilSsfDate0
            tlSsfSrchKey2.iDate(1) = ilSsfDate1
            ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
            If (ilRet = BTRV_ERR_NONE) And (blSport) Then
                Do
                    ilRet = gSSFGetPrevious(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If (ilRet = BTRV_ERR_NONE) Then
                        If (tmSsf.iVefCode <> ilVehCode) Or (tmSsf.iDate(0) <> ilSsfDate0) Or (tmSsf.iDate(1) <> ilSsfDate1) Then
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop
                ilRet = gSSFGetNext(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                ilType = tmSsf.iType
            End If
            'If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.sType <> slType) Or (tmSsf.iVefCode <> ilVehCode) Or (tmSsf.iDate(0) <> ilSsfDate0) Or (tmSsf.iDate(1) <> ilSsfDate1) Then
            If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> ilVehCode) Or (tmSsf.iDate(0) <> ilSsfDate0) Or (tmSsf.iDate(1) <> ilSsfDate1) Then
                'If airing- then use first Ssf prior to date defined
                If tmVef.sType = "A" Then
                    ilSsfDate0 = 0
                    ilSsfDate1 = 0
                    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    tlSsfSrchKey.iType = ilType
                    tlSsfSrchKey.iVefCode = ilVehCode
                    tlSsfSrchKey.iDate(0) = ilLogDate0
                    tlSsfSrchKey.iDate(1) = ilLogDate1
                    tlSsfSrchKey.iStartTime(0) = 0
                    tlSsfSrchKey.iStartTime(1) = 6144   '24*256
                    ilRet = gSSFGetLessOrEqual(hmSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefCode = ilVehCode)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode)
                        gUnpackDate tmSsf.iDate(0), tmSsf.iDate(1), slSsfDate
                        If (ilDay = gWeekDayStr(slSsfDate)) And (tmSsf.iStartTime(0) = 0) And (tmSsf.iStartTime(1) = 0) Then
                            ilSsfDate0 = tmSsf.iDate(0)
                            ilSsfDate1 = tmSsf.iDate(1)
                            Exit Do
                        End If
                        ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                        ilRet = gSSFGetPrevious(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            End If
            DoEvents
            'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefCode = ilVehCode) Then
            If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode) Then
                gUnpackDate ilSsfDate0, ilSsfDate1, slSsfDate
                'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
                Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
                    'Loop thru Ssf and move records to export
                    ReDim slBus(0 To 0) As String
                    slBus(0) = ""
                    If blSport Then
                        tmGsfSrchKey3.iVefCode = ilVefCode
                        tmGsfSrchKey3.iGameNo = ilType
                        ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                        Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode) And (tmGsf.iGameNo = ilType)
                            If (tmGsf.iAirDate(0) = ilSsfDate0) And (tmGsf.iAirDate(1) = ilSsfDate1) Then
                                slBuses = Trim$(tmGsf.sBus)
                                slBus = Split(slBuses, ";")
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        Loop
                    End If
                    ilEvt = 1
                    Do While ilEvt <= tmSsf.iCount
                       LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        If (tmProg.iRecType = 1) Or ((tmProg.iRecType >= 2) And (tmProg.iRecType <= 9)) Then
                            gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llEvtTime
                            If llEvtTime > llEndTime Then
                                ilWithinTime = False
                                Exit Do
                            End If
                            If llEvtTime >= llStartTime Then
                                ilWithinTime = True
                            End If
                        End If
                        '4/16/14: Test avail automation attribute
                        blAvailOk = True
                        If (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then
                           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            ilAnf = gBinarySearchAnf(tmAvail.ianfCode, tgAvailAnf())
                            If ilAnf <> -1 Then
                                If tgAvailAnf(ilAnf).sAutomationExport = "N" Then
                                    blAvailOk = False
                                End If
                            End If
                            If blBypassZeroUnits Then
                                If (tmAvail.iAvInfo And &H1F <= 0) Or (tmAvail.iLen <= 0) Then
                                    blAvailOk = False
                                End If
                            End If
                        End If
                        ilEvtFdIndex = -1
                        If (ilWithinTime) And (blAvailOk) Then
                            If tmProg.iRecType = 1 Then    'Program
                                ilAirHour = tmProg.iStartTime(1) \ 256  'Obtain Hour
                                If (llDate > llEDate) And (ilAirHour >= 6) Then
                                    ''Done processing
                                    'mExpNYDay = True
                                    'Exit Function
                                End If
                            ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
                               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                ilAirHour = tmAvail.iTime(1) \ 256  'Obtain Hour
                                If (llDate > llEDate) And (ilAirHour >= 6) Then
                                    ''Done processing
                                    'mExpNYDay = True
                                    'Exit Function
                                End If
                                ilSubFeed = False
                                If (tmVef.sType = "A") Or (ilEgfFound) Then
                                    If llPrevAvailEndTime = -1 Then
                                        ilNoPasses = 1
                                    Else
                                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAvailTime
                                        If llAvailTime = llPrevAvailEndTime Then
                                            ilNoPasses = 2
                                        Else
                                            ilNoPasses = 1
                                        End If
                                    End If
                                    ilEgfExist = False
                                    For ilPass = 1 To ilNoPasses Step 1
                                        ilPassFound = ilPass
                                        'Obtain engineering entry to see if avail is sent
                                        tmEgfSrchKey.iVefCode = ilVehCode
                                        tmEgfSrchKey.sAirDay = slDay
                                        tmEgfSrchKey.iStartDate(0) = ilEgfDate0
                                        tmEgfSrchKey.iStartDate(1) = ilEgfDate1
                                        If ilPass = 1 Then
                                            ilEgfTime0 = tmAvail.iTime(0)
                                            ilEgfTime1 = tmAvail.iTime(1)
                                            tmEgfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                            tmEgfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                        Else
                                            ilEgfTime0 = tmPrevAvail.iTime(0)
                                            ilEgfTime1 = tmPrevAvail.iTime(1)
                                            tmEgfSrchKey.iAirTime(0) = tmPrevAvail.iTime(0)
                                            tmEgfSrchKey.iAirTime(1) = tmPrevAvail.iTime(1)
                                        End If
                                        ilRet = btrGetEqual(hmEgf, tmEgf, imEgfRecLen, tmEgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmEgf.iVefCode = ilVehCode) And (tmEgf.sAirDay = slDay) And (tmEgf.iStartDate(0) = ilEgfDate0) And (tmEgf.iStartDate(1) = ilEgfDate1) And (tmEgf.iAirTime(0) = ilEgfTime0) And (tmEgf.iAirTime(1) = ilEgfTime1)
                                            ilEgfExist = True
                                            ilTerminated = False
                                            'If (tmEgf.sFed = "N") Or (tmEgf.sZone <> "CST") Then
                                            If (tmEgf.sFed = "N") Then
                                                ilTerminated = True
                                            Else
                                                If (tmEgf.iTermDate(1) <> 0) Or (tmEgf.iTermDate(0) <> 0) Then
                                                    If (tmEgf.iTermDate(1) < tmEgf.iStartDate(1)) Or ((tmEgf.iStartDate(1) = tmEgf.iTermDate(1)) And (tmEgf.iTermDate(0) < tmEgf.iStartDate(0))) Then
                                                        ilTerminated = True
                                                    End If
                                                End If
                                            End If
                                            If Not ilTerminated Then
                                                If tmEgf.iMnfSubFeed > 0 Then
                                                    ilSubFeed = True
                                                    Exit Do
                                                End If
                                            End If
                                            ilRet = btrGetNext(hmEgf, tmEgf, imEgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Loop
                                        If ilEgfExist Then
                                            Exit For
                                        End If
                                    Next ilPass
                                Else
                                    ilEgfExist = True   'Set to allow vehciles without Egf to be processed
                                End If
                                If Not ilEgfExist Then
                                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                    Screen.MousePointer = vbDefault
                                    ilRet = MsgBox(smVehName & " missing Engineering link at " & slDate & ": " & slTime, vbOKCancel + vbQuestion + vbApplicationModal, "Find Error")
                                    Screen.MousePointer = vbHourglass
                                    If ilRet = 2 Then   'Cancel
                                        mExpNYDay = False
                                        Exit Function
                                    End If
                                Else
                                    'Loop on spots, then add conflicting spots
                                    If (tmVef.sType = "A") Then
                                        tmVlfSrchKey1.iAirCode = ilVehCode
                                        tmVlfSrchKey1.iAirDay = Val(slDay)
                                        tmVlfSrchKey1.iEffDate(0) = ilVlfDate0
                                        tmVlfSrchKey1.iEffDate(1) = ilVlfDate1
                                        tmVlfSrchKey1.iAirTime(0) = tmAvail.iTime(0)
                                        tmVlfSrchKey1.iAirTime(1) = tmAvail.iTime(1)
                                        tmVlfSrchKey1.iAirPosNo = 0
                                        tmVlfSrchKey1.iAirSeq = 1
                                        ilRet = btrGetGreaterOrEqual(hmVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) And (tmVlf.iEffDate(0) = ilVlfDate0) And (tmVlf.iEffDate(1) = ilVlfDate1) And (tmVlf.iAirTime(0) = tmAvail.iTime(0)) And (tmVlf.iAirTime(1) = tmAvail.iTime(1))
                                            ilTerminated = False
                                            If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                                If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                                    ilTerminated = True
                                                End If
                                            End If
                                            If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                                'If (tmCTSsf.sType <> slType) Or (tmCTSsf.iVefCode <> tmVlf.iSellCode) Or (tmCTSsf.iDate(0) <> ilLogDate0) Or (tmCTSsf.iDate(1) <> ilLogDate1) Then
                                                If (tmCTSsf.iType <> ilType) Or (tmCTSsf.iVefCode <> tmVlf.iSellCode) Or (tmCTSsf.iDate(0) <> ilLogDate0) Or (tmCTSsf.iDate(1) <> ilLogDate1) Then
                                                    tmVefSrchKey.iCode = tmVlf.iSellCode
                                                    ilRet = btrGetEqual(hmVef, tmSVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                    If (ilRet = BTRV_ERR_NONE) Then
                                                        ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                        tlSsfSrchKey.iType = ilType
                                                        tlSsfSrchKey.iVefCode = tmVlf.iSellCode
                                                        tlSsfSrchKey.iDate(0) = ilLogDate0
                                                        tlSsfSrchKey.iDate(1) = ilLogDate1
                                                        tlSsfSrchKey.iStartTime(0) = 0
                                                        tlSsfSrchKey.iStartTime(1) = 0
                                                        ilRet = gSSFGetEqual(hmCTSsf, tmCTSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                    End If
                                                End If
                                                Do While (ilRet = BTRV_ERR_NONE) And (tmCTSsf.iType = ilType) And (tmCTSsf.iVefCode = tmVlf.iSellCode) And (tmCTSsf.iDate(0) = ilLogDate0) And (tmCTSsf.iDate(1) = ilLogDate1)
                                                    For ilSIndex = 1 To tmCTSsf.iCount Step 1
                                                        tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSIndex)
                                                        If ((tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9)) Then
                                                            If (tmAvailTest.iTime(0) = tmVlf.iSellTime(0)) And (tmAvailTest.iTime(1) = tmVlf.iSellTime(1)) Then
                                                                'Determine if any unsold time
                                                                ilAvailUnits = tmAvailTest.iAvInfo And &H1F
                                                                ilAvailLen = tmAvailTest.iLen
                                                                ilSpotLen = 0
                                                                For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                                   LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                    If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                                                        ilSpotLen = ilSpotLen + (tmSpot.iPosLen And &HFFF)
                                                                        ilSpotUnits = ilSpotUnits + 1
                                                                    End If
                                                                Next ilSpot
                                                                If ilAvailLen > ilSpotLen Then
                                                                    tmEgfSrchKey.sAirDay = slDay
                                                                    tmEgfSrchKey.iStartDate(0) = ilEgfDate0
                                                                    tmEgfSrchKey.iStartDate(1) = ilEgfDate1
                                                                    tmEgfSrchKey.iAirTime(0) = ilEgfTime0   'tmAvail.iTime(0)
                                                                    tmEgfSrchKey.iAirTime(1) = ilEgfTime1   'tmAvail.iTime(1)
                                                                    ilRet = btrGetEqual(hmEgf, tmEgf, imEgfRecLen, tmEgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmEgf.iVefCode = ilVehCode) And (tmEgf.sAirDay = slDay) And (tmEgf.iStartDate(0) = ilEgfDate0) And (tmEgf.iStartDate(1) = ilEgfDate1) And (tmEgf.iAirTime(0) = ilEgfTime0) And (tmEgf.iAirTime(1) = ilEgfTime1)
                                                                        ilTerminated = False
                                                                        'If (tmEgf.sFed = "N") Or (StrComp(Trim$(tmEgf.sZone), "CST", 1) <> 0) Then
                                                                        If (tmEgf.sFed = "N") Then
                                                                            ilTerminated = True
                                                                        Else
                                                                            If (tmEgf.iTermDate(1) <> 0) Or (tmEgf.iTermDate(0) <> 0) Then
                                                                                If (tmEgf.iTermDate(1) < tmEgf.iStartDate(1)) Or ((tmEgf.iStartDate(1) = tmEgf.iTermDate(1)) And (tmEgf.iTermDate(0) < tmEgf.iStartDate(0))) Then
                                                                                    ilTerminated = True
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        If Not ilTerminated Then
                                                                            If ilSubFeed Then
                                                                                If tmEgf.iMnfSubFeed > 0 Then
                                                                                    '6/1/16: Replace gosub
                                                                                    'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                                    mProcAdjDate slAirDate, llDate, ilAirHour, ilFeedHour
                                                                                    If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                                        slMsg = "Unsold: " & Trim$(tmSVef.sName)
                                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                        slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                                        tmAnfSrchKey.iCode = tmAvailTest.ianfCode
                                                                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                        If (ilRet = BTRV_ERR_NONE) Then
                                                                                            slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                                        End If
                                                                                        'Print #hmMsg, slMsg
                                                                                        gAutomationAlertAndLogHandler slMsg
                                                                                        'lbcMsg.AddItem slMsg
                                                                                        mAddMsgToListBox slMsg
                                                                                        Exit Do
                                                                                    End If
                                                                                    DoEvents
                                                                                    If (ilRet <> 0) Or (imTerminate = True) Then
                                                                                        mExpNYDay = False
                                                                                        Exit Function
                                                                                    End If
                                                                                End If
                                                                            Else
                                                                                If tmEgf.iMnfSubFeed = 0 Then
                                                                                    '6/1/16: Replace gosub
                                                                                    'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                                    mProcAdjDate slAirDate, llDate, ilAirHour, ilFeedHour
                                                                                    If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                                        slMsg = "Unsold: " & Trim$(tmSVef.sName)
                                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                        slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                                        tmAnfSrchKey.iCode = tmAvailTest.ianfCode
                                                                                        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                        If (ilRet = BTRV_ERR_NONE) Then
                                                                                            slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                                        End If
                                                                                        'Print #hmMsg, slMsg
                                                                                        gAutomationAlertAndLogHandler slMsg
                                                                                        'lbcMsg.AddItem slMsg
                                                                                        mAddMsgToListBox slMsg
                                                                                        Exit Do
                                                                                    End If
                                                                                    DoEvents
                                                                                    If (ilRet <> 0) Or (imTerminate = True) Then
                                                                                        mExpNYDay = False
                                                                                        Exit Function
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        ilRet = btrGetNext(hmEgf, tmEgf, imEgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                    Loop
                                                                End If
                                                                For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                                   LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                                    ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                    If ilRet = BTRV_ERR_NONE Then
                                                                        tmChfSrchKey.lCode = tmSdf.lChfCode
                                                                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                        If ilRet = BTRV_ERR_NONE Then
                                                                            tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                            If ilRet = BTRV_ERR_NONE Then
                                                                                slMsg = "Copy Missing: " & Trim$(tmSVef.sName)
                                                                                gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                                slMsg = slMsg & " for " & Trim$(str$(tmChf.lCntrNo)) & " " & Trim$(tmAdf.sName)
                                                                                tmEgfSrchKey.iVefCode = ilVehCode
                                                                                tmEgfSrchKey.sAirDay = slDay
                                                                                tmEgfSrchKey.iStartDate(0) = ilEgfDate0
                                                                                tmEgfSrchKey.iStartDate(1) = ilEgfDate1
                                                                                tmEgfSrchKey.iAirTime(0) = ilEgfTime0   'tmAvail.iTime(0)
                                                                                tmEgfSrchKey.iAirTime(1) = ilEgfTime1   'tmAvail.iTime(1)
                                                                                ilRet = btrGetEqual(hmEgf, tmEgf, imEgfRecLen, tmEgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                Do While (ilRet = BTRV_ERR_NONE) And (tmEgf.iVefCode = ilVehCode) And (tmEgf.sAirDay = slDay) And (tmEgf.iStartDate(0) = ilEgfDate0) And (tmEgf.iStartDate(1) = ilEgfDate1) And (tmEgf.iAirTime(0) = ilEgfTime0) And (tmEgf.iAirTime(1) = ilEgfTime1)
                                                                                    ilTerminated = False
                                                                                    'If (tmEgf.sFed = "N") Or (StrComp(Trim$(tmEgf.sZone), "CST", 1) <> 0) Then
                                                                                    If (tmEgf.sFed = "N") Then
                                                                                        ilTerminated = True
                                                                                    Else
                                                                                        If (tmEgf.iTermDate(1) <> 0) Or (tmEgf.iTermDate(0) <> 0) Then
                                                                                            If (tmEgf.iTermDate(1) < tmEgf.iStartDate(1)) Or ((tmEgf.iStartDate(1) = tmEgf.iTermDate(1)) And (tmEgf.iTermDate(0) < tmEgf.iStartDate(0))) Then
                                                                                                ilTerminated = True
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                    If Not ilTerminated Then
                                                                                        If ilSubFeed Then
                                                                                            If tmEgf.iMnfSubFeed > 0 Then
                                                                                                '6/1/16: Repcaed GoSub
                                                                                                'GoSub lProcSpot
                                                                                                mProcSpot ilVefCode, ilVehCode, ilVpfIndex, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilFeedHour, slCodeStn, ilPrefeed, llSpotTime, ilCopy, slZone, slISCI, llCopyMissingSdfCode, slShortTitle, slRecord, ilRet, slMsg
                                                                                                DoEvents
                                                                                                If (ilRet <> 0) Or (imTerminate = True) Then
                                                                                                    mExpNYDay = False
                                                                                                    Exit Function
                                                                                                End If
                                                                                            End If
                                                                                        Else
                                                                                            If tmEgf.iMnfSubFeed = 0 Then
                                                                                                '6/1/16: Repcaed GoSub
                                                                                                'GoSub lProcSpot
                                                                                                mProcSpot ilVefCode, ilVehCode, ilVpfIndex, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilFeedHour, slCodeStn, ilPrefeed, llSpotTime, ilCopy, slZone, slISCI, llCopyMissingSdfCode, slShortTitle, slRecord, ilRet, slMsg
                                                                                                DoEvents
                                                                                                If (ilRet <> 0) Or (imTerminate = True) Then
                                                                                                    mExpNYDay = False
                                                                                                    Exit Function
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                    ilRet = btrGetNext(hmEgf, tmEgf, imEgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                                Loop
                                                                            Else
                                                                                gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                                Screen.MousePointer = vbDefault
                                                                                ''MsgBox Trim$(tmSVef.sName) & " advertiser get error at air date: time " & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                                gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " advertiser get error at air date: time " & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                                Screen.MousePointer = vbHourglass
                                                                                mExpNYDay = False
                                                                                Exit Function
                                                                            End If
                                                                        Else
                                                                            gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                            Screen.MousePointer = vbDefault
                                                                            ''MsgBox Trim$(tmSVef.sName) & " contract get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                            gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " contract get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                            Screen.MousePointer = vbHourglass
                                                                            mExpNYDay = False
                                                                            Exit Function
                                                                        End If
                                                                    Else
                                                                        gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                                                                        Screen.MousePointer = vbDefault
                                                                        ''MsgBox Trim$(tmSVef.sName) & " spot get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                        gAutomationAlertAndLogHandler Trim$(tmSVef.sName) & " spot get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                                        Screen.MousePointer = vbHourglass
                                                                        mExpNYDay = False
                                                                        Exit Function
                                                                    End If
                                                                Next ilSpot
                                                                Exit Do
                                                            End If
                                                        End If
                                                    Next ilSIndex
                                                    ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                    ilRet = gSSFGetNext(hmCTSsf, tmCTSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                Loop
                                            End If
                                            ilRet = btrGetNext(hmVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Loop
                                    Else
                                        ilAvailUnits = tmAvail.iAvInfo And &H1F
                                        ilAvailLen = tmAvail.iLen
                                        ilSpotLen = 0
                                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilSpot + ilEvt)
                                            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                                ilSpotLen = ilSpotLen + (tmSpot.iPosLen And &HFFF)
                                                ilSpotUnits = ilSpotUnits + 1
                                            End If
                                        Next ilSpot
                                        If ilAvailLen > ilSpotLen Then
                                            'Test if Avail is shown
                                            gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                            If ilEgfFound Then
                                                'Obtain engineering entry to see is avail is sent
                                                tmEgfSrchKey.iVefCode = ilVehCode
                                                tmEgfSrchKey.sAirDay = slDay
                                                tmEgfSrchKey.iStartDate(0) = ilEgfDate0
                                                tmEgfSrchKey.iStartDate(1) = ilEgfDate1
                                                tmEgfSrchKey.iAirTime(0) = ilEgfTime0   'tmAvail.iTime(0)
                                                tmEgfSrchKey.iAirTime(1) = ilEgfTime1   'tmAvail.iTime(1)
                                                ilRet = btrGetEqual(hmEgf, tmEgf, imEgfRecLen, tmEgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                Do While (ilRet = BTRV_ERR_NONE) And (tmEgf.iVefCode = ilVehCode) And (tmEgf.sAirDay = slDay) And (tmEgf.iStartDate(0) = ilEgfDate0) And (tmEgf.iStartDate(1) = ilEgfDate1) And (tmEgf.iAirTime(0) = ilEgfTime0) And (tmEgf.iAirTime(1) = ilEgfTime1)
                                                    ilTerminated = False
                                                    'If (tmEgf.sFed = "N") Or (StrComp(Trim$(tmEgf.sZone), "CST", 1) <> 0) Then
                                                    If (tmEgf.sFed = "N") Then
                                                        ilTerminated = True
                                                    Else
                                                        If (tmEgf.iTermDate(1) <> 0) Or (tmEgf.iTermDate(0) <> 0) Then
                                                            If (tmEgf.iTermDate(1) < tmEgf.iStartDate(1)) Or ((tmEgf.iStartDate(1) = tmEgf.iTermDate(1)) And (tmEgf.iTermDate(0) < tmEgf.iStartDate(0))) Then
                                                                ilTerminated = True
                                                            End If
                                                        End If
                                                    End If
                                                    If Not ilTerminated Then
                                                        If ilSubFeed Then
                                                            If tmEgf.iMnfSubFeed > 0 Then
                                                                '6/1/16: Replace gosub
                                                                'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                mProcAdjDate slAirDate, llDate, ilAirHour, ilFeedHour
                                                                If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                    slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                                    slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                    tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                    If (ilRet = BTRV_ERR_NONE) Then
                                                                        slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                    End If
                                                                    'Print #hmMsg, slMsg
                                                                    gAutomationAlertAndLogHandler slMsg
                                                                    'lbcMsg.AddItem slMsg
                                                                    mAddMsgToListBox slMsg
                                                                    Exit Do
                                                                End If
                                                                DoEvents
                                                                If (ilRet <> 0) Or (imTerminate = True) Then
                                                                    mExpNYDay = False
                                                                    Exit Function
                                                                End If
                                                            End If
                                                        Else
                                                            If tmEgf.iMnfSubFeed = 0 Then
                                                                '6/1/16: Replace gosub
                                                                'GoSub lProcAdjDate  'Result stored into slAirDate
                                                                mProcAdjDate slAirDate, llDate, ilAirHour, ilFeedHour
                                                                If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                                    slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                                    slMsg = slMsg & " on " & slDate & " at " & slTime
                                                                    tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                    If (ilRet = BTRV_ERR_NONE) Then
                                                                        slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                                    End If
                                                                    'Print #hmMsg, slMsg
                                                                    gAutomationAlertAndLogHandler slMsg
                                                                    'lbcMsg.AddItem slMsg
                                                                    mAddMsgToListBox slMsg
                                                                    Exit Do
                                                                End If
                                                                DoEvents
                                                                If (ilRet <> 0) Or (imTerminate = True) Then
                                                                    mExpNYDay = False
                                                                    Exit Function
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    ilRet = btrGetNext(hmEgf, tmEgf, imEgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                Loop
                                            Else
                                                tmEgf.iAirTime(0) = tmAvail.iTime(0)
                                                tmEgf.iAirTime(1) = tmAvail.iTime(1)
                                                tmEgf.iLocalTime(0) = tmAvail.iTime(0)
                                                tmEgf.iLocalTime(1) = tmAvail.iTime(1)
                                                tmEgf.iFeedTime(0) = tmAvail.iTime(0)
                                                tmEgf.iFeedTime(1) = tmAvail.iTime(1)
                                                tmEgf.sZone = ""
                                                tmEgf.iEtfCode = 0
                                                tmEgf.iEnfCode = 0
                                                tmEgf.sProgCode = ""
                                                tmEgf.iMnfFeed = 0
                                                tmEgf.sBus = ""
                                                tmEgf.sSchedule = ""
                                                tmEgf.iMnfSubFeed = 0
                                                '6/1/16: Replace gosub
                                                'GoSub lProcAdjDate  'Result stored into slAirDate
                                                mProcAdjDate slAirDate, llDate, ilAirHour, ilFeedHour
                                                If (gDateValue(slAirDate) >= llSDate) And (gDateValue(slAirDate) <= llEDate) Then
                                                    slMsg = "Unsold: " & Trim$(tmVef.sName)
                                                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                    slMsg = slMsg & " on " & slDate & " at " & slTime
                                                    tmAnfSrchKey.iCode = tmAvail.ianfCode
                                                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                    If (ilRet = BTRV_ERR_NONE) Then
                                                        slMsg = slMsg & " " & Trim$(tmAnf.sName)
                                                    End If
                                                    'Print #hmMsg, slMsg
                                                    gAutomationAlertAndLogHandler slMsg
                                                    'lbcMsg.AddItem slMsg
                                                    mAddMsgToListBox slMsg
                                                End If
                                            End If
                                        End If
                                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                            ilEvt = ilEvt + 1
                                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet = BTRV_ERR_NONE Then
                                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet = BTRV_ERR_NONE Then
                                                    tmAdfSrchKey.iCode = tmChf.iAdfCode
                                                    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        slMsg = "Copy Missing: " & Trim$(smVehName)
                                                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                        slMsg = slMsg & " on " & slDate & " at " & slTime
                                                        slMsg = slMsg & " for " & Trim$(str$(tmChf.lCntrNo)) & " " & Trim$(tmAdf.sName)
                                                        If ilEgfFound Then
                                                            'Obtain engineering entry to see is avail is sent
                                                            tmEgfSrchKey.iVefCode = ilVehCode
                                                            tmEgfSrchKey.sAirDay = slDay
                                                            tmEgfSrchKey.iStartDate(0) = ilEgfDate0
                                                            tmEgfSrchKey.iStartDate(1) = ilEgfDate1
                                                            tmEgfSrchKey.iAirTime(0) = ilEgfTime0   'tmAvail.iTime(0)
                                                            tmEgfSrchKey.iAirTime(1) = ilEgfTime1   'tmAvail.iTime(1)
                                                            ilRet = btrGetEqual(hmEgf, tmEgf, imEgfRecLen, tmEgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                            Do While (ilRet = BTRV_ERR_NONE) And (tmEgf.iVefCode = ilVehCode) And (tmEgf.sAirDay = slDay) And (tmEgf.iStartDate(0) = ilEgfDate0) And (tmEgf.iStartDate(1) = ilEgfDate1) And (tmEgf.iAirTime(0) = ilEgfTime0) And (tmEgf.iAirTime(1) = ilEgfTime1)
                                                                ilTerminated = False
                                                                'If (tmEgf.sFed = "N") Or (StrComp(Trim$(tmEgf.sZone), "CST", 1) <> 0) Then
                                                                If (tmEgf.sFed = "N") Then
                                                                    ilTerminated = True
                                                                Else
                                                                    If (tmEgf.iTermDate(1) <> 0) Or (tmEgf.iTermDate(0) <> 0) Then
                                                                        If (tmEgf.iTermDate(1) < tmEgf.iStartDate(1)) Or ((tmEgf.iStartDate(1) = tmEgf.iTermDate(1)) And (tmEgf.iTermDate(0) < tmEgf.iStartDate(0))) Then
                                                                            ilTerminated = True
                                                                        End If
                                                                    End If
                                                                End If
                                                                If Not ilTerminated Then
                                                                    If ilSubFeed Then
                                                                        If tmEgf.iMnfSubFeed > 0 Then
                                                                            '6/1/16: Repcaed GoSub
                                                                            'GoSub lProcSpot
                                                                            mProcSpot ilVefCode, ilVehCode, ilVpfIndex, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilFeedHour, slCodeStn, ilPrefeed, llSpotTime, ilCopy, slZone, slISCI, llCopyMissingSdfCode, slShortTitle, slRecord, ilRet, slMsg
                                                                            DoEvents
                                                                            If (ilRet <> 0) Or (imTerminate = True) Then
                                                                                mExpNYDay = False
                                                                                Exit Function
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        If tmEgf.iMnfSubFeed = 0 Then
                                                                            '6/1/16: Repcaed GoSub
                                                                            'GoSub lProcSpot
                                                                            mProcSpot ilVefCode, ilVehCode, ilVpfIndex, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilFeedHour, slCodeStn, ilPrefeed, llSpotTime, ilCopy, slZone, slISCI, llCopyMissingSdfCode, slShortTitle, slRecord, ilRet, slMsg
                                                                            DoEvents
                                                                            If (ilRet <> 0) Or (imTerminate = True) Then
                                                                                mExpNYDay = False
                                                                                Exit Function
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                                ilRet = btrGetNext(hmEgf, tmEgf, imEgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                            Loop
                                                        Else
                                                            For ilBus = 0 To UBound(slBus) Step 1
                                                                tmEgf.iAirTime(0) = tmAvail.iTime(0)
                                                                tmEgf.iAirTime(1) = tmAvail.iTime(1)
                                                                tmEgf.iLocalTime(0) = tmAvail.iTime(0)
                                                                tmEgf.iLocalTime(1) = tmAvail.iTime(1)
                                                                tmEgf.iFeedTime(0) = tmAvail.iTime(0)
                                                                tmEgf.iFeedTime(1) = tmAvail.iTime(1)
                                                                tmEgf.sZone = ""
                                                                tmEgf.iEtfCode = 0
                                                                tmEgf.iEnfCode = 0
                                                                tmEgf.sProgCode = ""
                                                                tmEgf.iMnfFeed = 0
                                                                tmEgf.sBus = slBus(ilBus)   '""
                                                                tmEgf.sSchedule = ""
                                                                tmEgf.iMnfSubFeed = 0
                                                                '6/1/16: Repcaed GoSub
                                                                'GoSub lProcSpot
                                                                mProcSpot ilVefCode, ilVehCode, ilVpfIndex, llSDate, llEDate, slAirDate, llAirDate, llAvailTime, llDate, ilAirHour, ilFeedHour, slCodeStn, ilPrefeed, llSpotTime, ilCopy, slZone, slISCI, llCopyMissingSdfCode, slShortTitle, slRecord, ilRet, slMsg
                                                                If (ilRet <> 0) Or (imTerminate = True) Then
                                                                    mExpNYDay = False
                                                                    Exit Function
                                                                End If
                                                            Next ilBus
                                                        End If
                                                    Else
                                                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                        Screen.MousePointer = vbDefault
                                                        ''MsgBox smVehName & " advertiser get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                        gAutomationAlertAndLogHandler smVehName & " advertiser get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                        Screen.MousePointer = vbHourglass
                                                        mExpNYDay = False
                                                        Exit Function
                                                    End If
                                                Else
                                                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                    Screen.MousePointer = vbDefault
                                                    ''MsgBox smVehName & " contract get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                    gAutomationAlertAndLogHandler smVehName & " contract get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                    Screen.MousePointer = vbHourglass
                                                    mExpNYDay = False
                                                    Exit Function
                                                End If
                                            Else
                                                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                                Screen.MousePointer = vbDefault
                                                ''MsgBox smVehName & " spot get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                gAutomationAlertAndLogHandler smVehName & " spot get error at air date: time" & slDate & ": " & slTime, vbOkOnly + vbCritical + vbApplicationModal, "Write Error"
                                                Screen.MousePointer = vbHourglass
                                                mExpNYDay = False
                                                Exit Function
                                            End If
                                        Next ilSpot
                                    End If
                                    If ilPassFound = 1 Then
                                        tmPrevAvail = tmAvail
                                        'Compute end time of avail
                                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                                        slLen = Trim$(str$(tmAvail.iLen)) & "s"
                                        gAddTimeLength slTime, slLen, "A", "1", slTime, slXMid
                                        gPackTime slTime, ilTime0, ilTime1
                                        gUnpackTimeLong ilTime0, ilTime1, False, llPrevAvailEndTime
                                    End If
                                End If
                            End If
                        End If
                        ilEvt = ilEvt + 1
                    Loop
                    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hmSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If (ilRet = BTRV_ERR_NONE) And (blSport) Then
                        ilType = tmSsf.iType
                    End If
                Loop
            End If
        End If
        lgEndIndex = UBound(smNewLines)
        gBlackoutTest 0, hmCif, hmMcf, 0, 0, 0, 0, 0, 0, 0, 0, smNewLines(), hmMsg, lbcMsg
    'Next llDate
    Next ilPrefeed
    mExpNYDay = True
    Exit Function

    Return

    Return
'lProcSpot:
'    'Record type
'    slRecord = ""
'    'Test if spot is to be sent
'    'If tmEgf.iMnfSubFeed > 0 Then
'    '    tmMnfSrchKey.iCode = tmEgf.iMnfSubFeed
'    '    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
'    '    slCodeStn = Left$(tmMnf.sCodeStn, 5)
'    'Else
'        slCodeStn = Left$(tmVef.sCodeStn, 5)
'    'End If
'    '6/1/16: Replace gosub
'    'GoSub lProcAdjDate  'Result stored into slAirDate
'    mProcAdjDate slAirDate, llDate, ilAirHour, ilFeedHour
'    If tmPrefeedExpt(ilPrefeed).iFDay = -1 Then
'        If gDateValue(slAirDate) > llEDate Then
'            Return
'        End If
'        If gDateValue(slAirDate) < llSDate Then
'            Return
'        End If
'    End If
'    mSaveSpotTimes
'    '6/1/16: Replaced the GoSub
'    'GoSub lProcSpotTime 'Result stored into slSpotTime
'    mProcSpotTime slAirDate, llAirDate, llAvailTime, ilVehCode, slCodeStn, llSpotTime
'    If tmPrefeedExpt(ilPrefeed).iFDay <> -1 Then
'        If Not mPreFeed(llSDate, ilPrefeed, tmEgf.sZone, slAirDate, llSpotTime, tmEgf.sBus) Then
'            mRestoreSpotTimes
'            Return
'        End If
'        If gDateValue(slAirDate) > llEDate Then
'            mRestoreSpotTimes
'            Return
'        End If
'        If gDateValue(slAirDate) < llSDate Then
'            mRestoreSpotTimes
'            Return
'        End If
'    End If
'    'Get Copy for spot
'    slZone = Trim$(tmEgf.sZone)
'    If slZone = "" Then
'        slZone = "Oth"
'    End If
'
'    '8/29/14: Check for Airing vehicle copy (ttp 7064)
'    gGetAirCopy tmVef.sType, tmVef.iCode, ilVpfIndex, tmSdf, hmCrf, hmRsf, hmCvf, slZone
'
'    ilCopy = mObtainCopy(slZone)
'    If Not ilCopy Then
'        If llCopyMissingSdfCode <> tmSdf.lCode Then
'            Print #hmMsg, slMsg
'            'lbcMsg.AddItem slMsg
'            mAddMsgToListBox slMsg
'            llCopyMissingSdfCode = tmSdf.lCode
'        End If
'    End If
'    'Later if copy missing- output error and stop creation, delete file
'
'
'    'Air date
'    gObtainYearMonthDayStr slAirDate, True, slFYear, slFMonth, slFDay
'    slRecord = slRecord & right$(slFYear, 2) & slFMonth & slFDay & " "
'    ''Zone
'    'If tmEgf.sZone = "EST" Then
'    '    slRecord = slRecord & "03 "
'    'ElseIf tmEgf.sZone = "CST" Then
'    '    slRecord = slRecord & "04 "
'    'ElseIf tmEgf.sZone = "PST" Then
'    '    slRecord = slRecord & "06 "
'    'Else
'    '    slRecord = slRecord & "   "
'    'End If
'    'Schedule
'    slStr = Trim$(tmEgf.sSchedule)
'    Do While Len(slStr) < 2
'        slStr = "0" & slStr
'    Loop
'    slRecord = slRecord & slStr & " "
'    'Feed Time
'    llHour = llSpotTime \ 3600
'    llMin = llSpotTime Mod 3600
'    llSec = llMin Mod 60
'    llMin = llMin \ 60
'    slStr = Trim$(str$(llHour))
'    Do While Len(slStr) < 2
'        slStr = "0" & slStr
'    Loop
'    slRecord = slRecord & slStr
'    slStr = Trim$(str$(llMin))
'    Do While Len(slStr) < 2
'        slStr = "0" & slStr
'    Loop
'    slRecord = slRecord & slStr
'    slStr = Trim$(str$(llSec))
'    Do While Len(slStr) < 2
'        slStr = "0" & slStr
'    Loop
'    slRecord = slRecord & slStr & " "
'    'Bus
'    slRecord = slRecord & tmEgf.sBus & " "  'All 5 characters required
'    'Cart ID
'    If ilCopy Then
'        '12/26/13: Suppress Media code
'        '4/16/14: If Live, retain Media Code Name
'        'If (smEngrHideMediaCode = "Y") Then
'        If (smEngrHideMediaCode = "Y") And (Trim$(tmMcf.sName) <> "L") Then
'            slRecord = slRecord & tmCif.sName & " " 'retain 5 characters
'        Else
'            If Trim$(tmCif.sCut) = "" Then
'                slRecord = slRecord & Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
'            Else
'                slRecord = slRecord & Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
'            End If
'        End If
'    Else
'        slRecord = slRecord & "C    " & " "
'        imCopyMissing = True
'    End If
'    'Commercial Title
'    'If ilCopy And (Trim$(tmCpf.sName) <> "") Then
'    '    slShortTitle = UCase$(Left$(tmCpf.sName, 15))
'    'Else
'    '    slShortTitle = UCase$(Left$(tmChf.sProduct, 15))
'    'End If
'    If tmChf.iAdfCode <> tmAdf.iCode Then
'        tmAdfSrchKey.iCode = tmChf.iAdfCode
'        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    End If
'    slShortTitle = UCase$(Left$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf), 15))
'    slRecord = slRecord & slShortTitle & " "
'    'Spot length
'    If tmSdf.iLen < 60 Then
'        slStr = Trim$(str$(tmSdf.iLen))
'        Do While Len(slStr) < 4
'            slStr = "0" & slStr
'        Loop
'        slRecord = slRecord & slStr & " "
'    ElseIf tmSdf.iLen = 60 Then
'        slRecord = slRecord & "0100" & " "
'    Else
'        llMin = tmSdf.iLen
'        llSec = llMin Mod 60
'        llMin = llMin \ 60
'        slStr = Trim$(str$(llMin))
'        Do While Len(slStr) < 2
'            slStr = "0" & slStr
'        Loop
'        slRecord = slRecord & slStr
'        slStr = Trim$(str$(llSec))
'        Do While Len(slStr) < 2
'            slStr = "0" & slStr
'        Loop
'        slRecord = slRecord & slStr & " "
'    End If
'    'Vehicle code
'    slRecord = slRecord & slCodeStn
'    'Set to 52 so that EAS is correct
'    Do While Len(slRecord) < 52
'        slRecord = slRecord & " "
'    Loop
'    'ISCI
'    If (smEASOrASPOrESPN = "ASP") Or (smEASOrASPOrESPN = "ESPN") Then
'        slISCI = Trim$(tmCpf.sISCI)
'        slISCI = gFileNameFilter(slISCI)
'        Do While Len(slISCI) < 20
'            slISCI = slISCI & " "
'        Loop
'        slRecord = slRecord & slISCI
'    End If
'    ilRet = 0
'    smNewLines(UBound(smNewLines)) = slRecord
'    tgSpotSum(UBound(tgSpotSum)).lDate = gDateValue(slAirDate)
'    tgSpotSum(UBound(tgSpotSum)).iVefCode = ilVefCode
'    tgSpotSum(UBound(tgSpotSum)).iNewIndex = UBound(smNewLines)
'    tgSpotSum(UBound(tgSpotSum)).iMnfComp(0) = tmSpot.iMnfComp(0)
'    tgSpotSum(UBound(tgSpotSum)).iMnfComp(1) = tmSpot.iMnfComp(1)
'    tgSpotSum(UBound(tgSpotSum)).iAdfCode = tmSdf.iAdfCode
'    tgSpotSum(UBound(tgSpotSum)).lTime = llSpotTime
'    tgSpotSum(UBound(tgSpotSum)).iLen = tmSdf.iLen
'    tgSpotSum(UBound(tgSpotSum)).sShortTitle = slShortTitle
'    ReDim Preserve tgSpotSum(0 To UBound(tgSpotSum) + 1) As SPOTSUM
'    ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 72
'    'On Error GoTo mExpNYDayErr:
'    'Print #hmTo, slStoreRec
'    'On Error GoTo 0
'    Return
'lProcAdjDate:
'    'Test if Air time is AM and Feed Time is PM. If so, adjust date
'    slAirDate = Format$(llDate, "m/d/yy")
'    ilAirHour = tmEgf.iAirTime(1) \ 256  'Obtain Hour
'    ilFeedHour = tmEgf.iFeedTime(1) \ 256  'Obtain Hour
'    If (ilAirHour < 6) And (ilFeedHour > 17) Then
'        slAirDate = gDecOneDay(slAirDate)
'    ElseIf (ilAirHour > 17) And (ilFeedHour < 6) Then
'        slAirDate = gIncOneDay(slAirDate)
'    End If
'    Return
'lProcSpotTime:
'    llAirDate = gDateValue(slAirDate)
'    gUnpackTimeLong tmEgf.iFeedTime(0), tmEgf.iFeedTime(1), False, llAvailTime
'    For ilTest = LBound(tmSpotTimes) To UBound(tmSpotTimes) - 1 Step 1
'        If (tmSpotTimes(ilTest).iVefCode = ilVehCode) And (tmSpotTimes(ilTest).sCodeStn = slCodeStn) And (tmSpotTimes(ilTest).sBus = tmEgf.sBus) And (tmSpotTimes(ilTest).lAirDate = llAirDate) And (tmSpotTimes(ilTest).lAvailTime = llAvailTime) Then
'            llSpotTime = tmSpotTimes(ilTest).lNextSpotTime
'            tmSpotTimes(ilTest).lNextSpotTime = tmSpotTimes(ilTest).lNextSpotTime + tmSdf.iLen
'            Return
'        End If
'    Next ilTest
'    ReDim Preserve tmSpotTimes(0 To UBound(tmSpotTimes) + 1) As SPOTTIMESNY
'    tmSpotTimes(UBound(tmSpotTimes) - 1).iVefCode = ilVehCode
'    tmSpotTimes(UBound(tmSpotTimes) - 1).sCodeStn = slCodeStn
'    tmSpotTimes(UBound(tmSpotTimes) - 1).sBus = tmEgf.sBus
'    tmSpotTimes(UBound(tmSpotTimes) - 1).lAirDate = llAirDate
'    tmSpotTimes(UBound(tmSpotTimes) - 1).lAvailTime = llAvailTime
'    tmSpotTimes(UBound(tmSpotTimes) - 1).lNextSpotTime = llAvailTime + tmSdf.iLen
'    llSpotTime = llAvailTime
'    Return
End Function
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
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilTest As Integer
    Dim ilVefCode As Integer
    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    imLBCDCtrls = 1
    imExporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0

    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmCpf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imAnfRecLen = Len(tmAnf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)
    hmSif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imSifRecLen = Len(tmSif)
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)
    hmCif = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmTzf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imTzfRecLen = Len(tmTzf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmVlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)
    hmEgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmEgf, "", sgDBPath & "Egf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imEgfRecLen = Len(tmEgf)
    hmPff = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPff, "", sgDBPath & "Pff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imPffRecLen = Len(tmPff)
    hmPbf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPbf, "", sgDBPath & "Pbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imPbfRecLen = Len(tmPbf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)
    hmBof = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imBofRecLen = Len(tmBof)
    hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    imGsfRecLen = Len(tmGsf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    hmCTSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCTSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    
    '8/29/14: Check for Airing vehicle copy (ttp 7064)
    hmCrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    hmRsf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    hmCvf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ExpNY
    On Error GoTo 0
    
    'Populate arrays to determine if records exist
    mVehPop
    If imTerminate Then
        Screen.MousePointer = vbDefault
        'mTerminate
        Exit Sub
    End If
    ilRet = gPopAdvtBox(ExpNY, lbcAdvt, tmAdvertiser(), smAdvertiserTag)
    ''Select Satellite Music; SMN Mix; Tom Joyner
    'For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
    '    slNameCode = tmUserVehicle(ilLoop).sKey    'Traffic!lbcUserVehicle.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    ilVefCode = Val(slCode)
    '    'For ilTest = 0 To UBound(tgVpf) Step 1
    '    '    If ilVefCode = tgVpf(ilTest).iVefKCode Then
    '        ilTest = gBinarySearchVpf(ilVefCode)
    '        If ilTest <> -1 Then
    '            If smEASOrASP = "ASP" Then
    '                If tgVpf(ilTest).sExpHiNYISCI = "Y" Then
    '                    lbcVehicle.Selected(ilLoop) = True
    '                End If
    '            Else
    '                If tgVpf(ilTest).sExpHiNY = "Y" Then
    '                    lbcVehicle.Selected(ilLoop) = True
    '                End If
    '            End If
    '    '        Exit For
    '        End If
    '    'Next ilTest
    ''    If StrComp(lbcVehicle.List(ilLoop), "ABC Business Report", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    ElseIf StrComp(lbcVehicle.List(ilLoop), "Bob Brinker", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    ElseIf StrComp(lbcVehicle.List(ilLoop), "Contemporary", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    ElseIf StrComp(lbcVehicle.List(ilLoop), "Direction", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    ElseIf StrComp(lbcVehicle.List(ilLoop), "Entertainment", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    ElseIf StrComp(lbcVehicle.List(ilLoop), "ESPN-Sportsbabe", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    ElseIf StrComp(lbcVehicle.List(ilLoop), "Jim Hightower", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    ElseIf StrComp(lbcVehicle.List(ilLoop), "Information", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    ElseIf StrComp(lbcVehicle.List(ilLoop), "Rock", 1) = 0 Then
    ''        lbcVehicle.Selected(ilLoop) = True
    ''    End If
    'Next ilLoop
    ilRet = gObtainSAF()
    smEngrHideMediaCode = "N"
    For ilLoop = 0 To UBound(tgSaf) - 1 Step 1
        If tgSaf(ilLoop).iVefCode <= 0 Then
            If (Asc(tgSaf(0).sFeatures1) And ENGRHIDEMEDIACODE) = ENGRHIDEMEDIACODE Then
                smEngrHideMediaCode = "Y"
            End If
            Exit For
        End If
    Next ilLoop
    '4/16/14: Add test of avail attribute
    ilRet = gObtainAvail()
    For ilLoop = LBound(imEvtType) To UBound(imEvtType) Step 1
        imEvtType(ilLoop) = True
    Next ilLoop
    imEvtType(0) = False 'Don't include library names
    'plcGauge.Move ExpNY.Width / 2 - plcGauge.Width / 2
    'cmcFileConv.Move ExpNY.Width / 2 - cmcFileConv.Width / 2
    'cmcCancel.Move ExpNY.Width / 2 - cmcCancel.Width / 2 - cmcCancel.Width
    'cmcReport.Move ExpNY.Width / 2 - cmcReport.Width / 2 + cmcReport.Width
    imBSMode = False
    imCalType = 0   'Standard
    mInitBox
    slStr = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(slStr)
    slStr = gObtainNextMonday(slStr)
    gObtainMonthYear imCalType, slStr, imCalMonth, imCalYear
    pbcCalendar_Paint   'mBoxCalDate called within paint
    lacDate.Visible = False
    gCenterStdAlone ExpNY
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    gAutomationAlertAndLogHandler ""
    gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
    
    Exit Sub
mInitErr:
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim ilLoop As Integer
    'Calendar
    For ilLoop = 1 To 7 Step 1
        gSetCtrl tmCDCtrls(ilLoop), 30 + 255 * (ilLoop - 1), 225, 240, fgBoxGridH
    Next ilLoop
    plcCalendar.Move edcStartDate.Left, edcStartDate.Top + edcStartDate.Height
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCopy                     *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Private Function mObtainCopy(slZone As String) As Integer
'
'   mObtainCopy
'       Where:
'           tmSdf(I)- Spot record
'           tmCif(O)- Inventory
'           tmCpf(O)- Product/ISCI record
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilCifFound As Integer
    ilCifFound = False
    tmCpf.sISCI = ""
    tmCpf.sName = ""
    tmCpf.sCreative = ""
    tmMcf.sName = "C"
    tmMcf.sPrefix = "C"
    If tmSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        tmCifSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilCifFound = True
        End If
    ElseIf tmSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tmSdf.sPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        tmTzfSrchKey.lCode = tmSdf.lCopyCode
        ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tmTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tmTzf.sZone(ilIndex - 1), slZone, 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                    ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilCifFound = True
                    End If
                    Exit For
                End If
            End If
        Next ilIndex
        If Not ilCifFound Then
            For ilIndex = 1 To 6 Step 1
                If tmTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                    If StrComp(tmTzf.sZone(ilIndex - 1), "Oth", 1) = 0 Then
                        ' Read CIF using lCopyCode from SDF
                        tmCifSrchKey.lCode = tmTzf.lCifZone(ilIndex - 1)
                        ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            ilCifFound = True
                        End If
                        Exit For
                    End If
                End If
            Next ilIndex
        End If
    End If
    If ilCifFound Then
        ' Read CPF using lCpfCode from CIF
        If tmCif.lcpfCode > 0 Then
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                tmCpf.sISCI = ""
                tmCpf.sName = ""
                tmCpf.sCreative = ""
            End If
        Else
            tmCpf.sISCI = ""
            tmCpf.sName = ""
            tmCpf.sCreative = ""
        End If
        tmMcfSrchKey.iCode = tmCif.iMcfCode
        ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmMcf.sName = "C"
            tmMcf.sPrefix = "C"
            mObtainCopy = False
            Exit Function
        End If
        mObtainCopy = True
        Exit Function
    End If
    mObtainCopy = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    'On Error GoTo mOpenMsgFileErr:
    'slToFile = sgExportPath & "ExpNY.Txt"
    slToFile = sgDBPath & "Messages\" & "ExpEngr.Txt"
    sgMessageFile = slToFile
    'slDateTime = FileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = lmNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, ""
    
    'Print #hmMsg, "Export Engineering: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    gAutomationAlertAndLogHandler "Export Engineering: " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
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
    slCommand = sgCommandStr  'Command$
    'If StrComp(slCommand, "Debug", 1) = 0 Then
    '    igStdAloneMode = True 'Switch from/to stand alone mode
    '    sgCallAppName = ""
    '    slStr = "Guide"
    '    ilTestSystem = False
    '    imShowHelpMsg = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            ''MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            gAutomationAlertAndLogHandler "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
        'imShowHelpMsg = True
        'ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
        'If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
        '    imShowHelpMsg = False
        'End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone ExpNY, slStr, ilTestSystem
    ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCmmCallSource = Val(slStr)
    ilRet = gParseItem(slCommand, 4, "\", smEASOrASPOrESPN)    'Get call source
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    'End If
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

    Screen.MousePointer = vbDefault
    igParentRestarted = False
    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExpNY
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim ilVefCode As Integer
    Dim slName As String
    Dim ilVff As Integer
    Dim ilVpf As Integer
    Dim blRetain As Boolean
    Dim ilLoop As Integer
    Dim slCode As String
    Dim llLen As Long
    
    'ilRet = gPopUserVehicleBox(ExpNY, VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcVehicle, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(ExpNY, VEHCONV_W_FEED + VEHAIRING + ACTIVEVEH, lbcVehicle, tmUserVehicle(), smUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ExpNY
        On Error GoTo 0
    End If
    ilIndex = UBound(tmUserVehicle) - 1
    Do
        slNameCode = tmUserVehicle(ilIndex).sKey    'Traffic!lbcUserVehicle.List(ilIndex)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
        blRetain = False
        'For ilVff = LBound(tgVff) To UBound(tgVff) - 1 Step 1
        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
            If ilVefCode = tgVff(ilVff).iVefCode Then
                ilVpf = gBinarySearchVpf(ilVefCode)
                If ilVpf <> -1 Then
                    If smEASOrASPOrESPN = "ASP" Then
                        If tgVpf(ilVpf).sExpHiNYISCI = "Y" Then
                            blRetain = True
                        End If
                    ElseIf smEASOrASPOrESPN = "ESPN" Then
                        If tgVff(ilVff).sExportNYESPN = "Y" Then
                            blRetain = True
                        End If
                    Else
                        If tgVpf(ilVpf).sExpHiNY = "Y" Then
                            blRetain = True
                        End If
                    End If
                    Exit For
                End If
            End If
        Next ilVff
        If Not blRetain Then
            For ilLoop = ilIndex To UBound(tmUserVehicle) - 1 Step 1
                tmUserVehicle(ilLoop) = tmUserVehicle(ilLoop + 1)
            Next ilLoop
            ReDim Preserve tmUserVehicle(LBound(tmUserVehicle) To UBound(tmUserVehicle) - 1) As SORTCODE
        End If
        ilIndex = ilIndex - 1
    Loop While ilIndex >= LBound(tmUserVehicle)
    
    llLen = 0
    lbcVehicle.Clear
    For ilLoop = 0 To UBound(tmUserVehicle) - 1 Step 1
        slNameCode = tmUserVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            Exit Sub
        End If
        ilRet = gParseItem(slName, 3, "|", slName)
        If ilRet <> CP_MSG_NONE Then
            Exit Sub
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcVehicle.AddItem slName  'Add ID to list box
    Next ilLoop
    'Select all
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        lbcVehicle.Selected(ilLoop) = True
    Next ilLoop
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llDate As Long
    Dim ilWkDay As Integer
    Dim ilRowNo As Integer
    Dim slDay As String
    ilRowNo = 0
    llDate = lmCalStartDate
    Do
        ilWkDay = gWeekDayLong(llDate)
        slDay = Trim$(str$(Day(llDate)))
        If (X >= tmCDCtrls(ilWkDay + 1).fBoxX) And (X <= (tmCDCtrls(ilWkDay + 1).fBoxX + tmCDCtrls(ilWkDay + 1).fBoxW)) Then
            If (Y >= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15)) And (Y <= tmCDCtrls(ilWkDay + 1).fBoxY + ilRowNo * (tmCDCtrls(ilWkDay + 1).fBoxH + 15) + tmCDCtrls(ilWkDay + 1).fBoxH) Then
                edcStartDate.Text = Format$(llDate, "m/d/yy")
                edcStartDate.SelStart = 0
                edcStartDate.SelLength = Len(edcStartDate.Text)
                imBypassFocus = True
                edcStartDate.SetFocus
                Exit Sub
            End If
        End If
        If ilWkDay = 6 Then
            ilRowNo = ilRowNo + 1
        End If
        llDate = llDate + 1
    Loop Until llDate > lmCalEndDate
    edcStartDate.SetFocus
End Sub
Private Sub pbcCalendar_Paint()
    Dim slStr As String
    slStr = Trim$(str$(imCalMonth)) & "/15/" & Trim$(str$(imCalYear))
    lacCalName.Caption = gMonthYearFormat(slStr)
    gPaintCalendar imCalMonth, imCalYear, imCalType, pbcCalendar, tmCDCtrls(), lmCalStartDate, lmCalEndDate
    mBoxCalDate
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    If smEASOrASPOrESPN = "ASP" Then
        plcScreen.Print "Export Engineering ASP Feed"
    ElseIf smEASOrASPOrESPN = "ESPN" Then
        plcScreen.Print "Export Engineering ESPN Feed"
    Else
        plcScreen.Print "Export Engineering EAS Feed"
    End If
End Sub

Private Sub mAddMsgToListBox(slMsg As String)
    Dim llValue As Long
    Dim llRg As Long
    Dim llRet As Long

    lbcMsg.AddItem slMsg
    If (Traffic.pbcArial.TextWidth(slMsg)) > lmMaxWidth Then
        lmMaxWidth = Traffic.pbcArial.TextWidth(slMsg)
        If lmMaxWidth > lbcMsg.Width Then
            llValue = lmMaxWidth / 15 + 120
            llRg = 0
            llRet = SendMessageByNum(lbcMsg.HWnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
        End If
    End If
End Sub




Private Sub mGetPBfRecs(llPffCode As Long)
    Dim ilRet As Integer
    Dim ilUpper As Integer
    
    ReDim tmPBFInfo(0 To 0) As PBFINFO
    If llPffCode <= 0 Then
        Exit Sub
    End If
    tmPbfSrchKey1.lCode = llPffCode
    ilRet = btrGetEqual(hmPbf, tmPbf, imPbfRecLen, tmPbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While ilRet = BTRV_ERR_NONE
        If tmPbf.lPffCode <> llPffCode Then
            Exit Do
        End If
        ilUpper = UBound(tmPBFInfo)
        tmPBFInfo(ilUpper).sKey = ""
        tmPBFInfo(ilUpper).tPbf = tmPbf
        ReDim Preserve tmPBFInfo(0 To ilUpper + 1) As PBFINFO
        ilUpper = ilUpper + 1
        ilRet = btrGetNext(hmPbf, tmPbf, imPbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Loop
End Sub

Private Function mPreFeed(llStartDate As Long, ilPrefeed As Integer, slZone As String, slAirDate As String, llSpotTime As Long, slBus As String) As Integer
    Dim ilDay As Integer
    Dim ilPbf As Integer
    
    mPreFeed = True
    If tmPrefeedExpt(ilPrefeed).iFDay = -1 Then
        Exit Function
    End If
    ilDay = gWeekDayStr(slAirDate)
    If tmPrefeedExpt(ilPrefeed).iFDay = ilDay Then
        If (llSpotTime >= tmPrefeedExpt(ilPrefeed).lFStartTime) And (llSpotTime <= tmPrefeedExpt(ilPrefeed).lFEndTime) Then
            If (Left$(slZone, 1) = tmPrefeedExpt(ilPrefeed).sFZone) Or (tmPrefeedExpt(ilPrefeed).sFZone = "A") Then
                'Test Bus
                For ilPbf = 0 To UBound(tmPBFInfo) - 1 Step 1
                    If tmPBFInfo(ilPbf).tPbf.sFromBus = slBus Then
                        slBus = tmPBFInfo(ilPbf).tPbf.sToBus
                        ''slAirDate = Format(llStartDate, "m/d/yy")
                        'slAirDate = Format(llStartDate - (tmPrefeedExpt(ilPrefeed).iFDay - tmPrefeedExpt(ilPrefeed).iTDay), "m/d/yy")
                        slAirDate = Format(llStartDate - (gWeekDayLong(llStartDate) - tmPrefeedExpt(ilPrefeed).iTDay), "m/d/yy")
                        llSpotTime = llSpotTime + tmPrefeedExpt(ilPrefeed).lAdjTime
                        If llSpotTime < 0 Then
                            slAirDate = gDecOneDay(slAirDate)
                            llSpotTime = 86400 + llSpotTime
                        ElseIf llSpotTime > 86400 Then
                            slAirDate = gIncOneDay(slAirDate)
                            llSpotTime = llSpotTime - 86400
                        End If
                        Exit Function
                    End If
                Next ilPbf
            End If
        End If
    End If
    mPreFeed = False
End Function

Private Sub mSaveSpotTimes()
    Dim ilLoop As Integer
    
    ReDim tmSvSpotTimes(0 To UBound(tmSpotTimes)) As SPOTTIMESNY
    For ilLoop = 0 To UBound(tmSpotTimes) - 1 Step 1
        tmSvSpotTimes(ilLoop) = tmSpotTimes(ilLoop)
    Next ilLoop

End Sub

Private Sub mRestoreSpotTimes()
    Dim ilLoop As Integer
    
    ReDim tmSpotTimes(0 To UBound(tmSvSpotTimes)) As SPOTTIMESNY
    For ilLoop = 0 To UBound(tmSvSpotTimes) - 1 Step 1
        tmSpotTimes(ilLoop) = tmSvSpotTimes(ilLoop)
    Next ilLoop
End Sub


Private Sub mProcAdjDate(slAirDate As String, llDate As Long, ilAirHour As Integer, ilFeedHour As Integer)
    slAirDate = Format$(llDate, "m/d/yy")
    ilAirHour = tmEgf.iAirTime(1) \ 256  'Obtain Hour
    ilFeedHour = tmEgf.iFeedTime(1) \ 256  'Obtain Hour
    If (ilAirHour < 6) And (ilFeedHour > 17) Then
        slAirDate = gDecOneDay(slAirDate)
    ElseIf (ilAirHour > 17) And (ilFeedHour < 6) Then
        slAirDate = gIncOneDay(slAirDate)
    End If
End Sub

Private Sub mProcSpotTime(slAirDate As String, llAirDate As Long, llAvailTime As Long, ilVehCode As Integer, slCodeStn As String, llSpotTime As Long)
    Dim ilTest As Integer
    
    llAirDate = gDateValue(slAirDate)
    gUnpackTimeLong tmEgf.iFeedTime(0), tmEgf.iFeedTime(1), False, llAvailTime
    For ilTest = LBound(tmSpotTimes) To UBound(tmSpotTimes) - 1 Step 1
        If (tmSpotTimes(ilTest).iVefCode = ilVehCode) And (tmSpotTimes(ilTest).sCodeStn = slCodeStn) And (tmSpotTimes(ilTest).sBus = tmEgf.sBus) And (tmSpotTimes(ilTest).lAirDate = llAirDate) And (tmSpotTimes(ilTest).lAvailTime = llAvailTime) Then
            llSpotTime = tmSpotTimes(ilTest).lNextSpotTime
            tmSpotTimes(ilTest).lNextSpotTime = tmSpotTimes(ilTest).lNextSpotTime + tmSdf.iLen
            'Return
            Exit Sub
        End If
    Next ilTest
    ReDim Preserve tmSpotTimes(0 To UBound(tmSpotTimes) + 1) As SPOTTIMESNY
    tmSpotTimes(UBound(tmSpotTimes) - 1).iVefCode = ilVehCode
    tmSpotTimes(UBound(tmSpotTimes) - 1).sCodeStn = slCodeStn
    tmSpotTimes(UBound(tmSpotTimes) - 1).sBus = tmEgf.sBus
    tmSpotTimes(UBound(tmSpotTimes) - 1).lAirDate = llAirDate
    tmSpotTimes(UBound(tmSpotTimes) - 1).lAvailTime = llAvailTime
    tmSpotTimes(UBound(tmSpotTimes) - 1).lNextSpotTime = llAvailTime + tmSdf.iLen
    llSpotTime = llAvailTime

End Sub

Private Sub mProcSpot(ilVefCode As Integer, ilVehCode As Integer, ilVpfIndex As Integer, llSDate As Long, llEDate As Long, slAirDate As String, llAirDate As Long, llAvailTime As Long, llDate As Long, ilAirHour As Integer, ilFeedHour As Integer, slCodeStn As String, ilPrefeed As Integer, llSpotTime As Long, ilCopy As Integer, slZone As String, slISCI As String, llCopyMissingSdfCode As Long, slShortTitle As String, slRecord As String, ilRet As Integer, slMsg As String)
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slStr As String
    'Dim slMsg As String
    Dim llHour As Long
    Dim llMin As Long
    Dim llSec As Long
    
    ilRet = 0
    'Record type
    slRecord = ""
    'Test if spot is to be sent
    'If tmEgf.iMnfSubFeed > 0 Then
    '    tmMnfSrchKey.iCode = tmEgf.iMnfSubFeed
    '    ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    '    slCodeStn = Left$(tmMnf.sCodeStn, 5)
    'Else
        '12/20/17: Changed to match definition of the field
        'slCodeStn = Left$(tmVef.sCodeStn, 5)
        slCodeStn = Left$(tmVef.sCodeStn, 2)
    'End If
    '6/1/16: Replace gosub
    'GoSub lProcAdjDate  'Result stored into slAirDate
    mProcAdjDate slAirDate, llDate, ilAirHour, ilFeedHour
    If tmPrefeedExpt(ilPrefeed).iFDay = -1 Then
        If gDateValue(slAirDate) > llEDate Then
            'Return
            Exit Sub
        End If
        If gDateValue(slAirDate) < llSDate Then
            'Return
            Exit Sub
        End If
    End If
    mSaveSpotTimes
    '6/1/16: Replaced the GoSub
    'GoSub lProcSpotTime 'Result stored into slSpotTime
    mProcSpotTime slAirDate, llAirDate, llAvailTime, ilVehCode, slCodeStn, llSpotTime
    If tmPrefeedExpt(ilPrefeed).iFDay <> -1 Then
        If Not mPreFeed(llSDate, ilPrefeed, tmEgf.sZone, slAirDate, llSpotTime, tmEgf.sBus) Then
            mRestoreSpotTimes
            'Return
            Exit Sub
        End If
        If gDateValue(slAirDate) > llEDate Then
            mRestoreSpotTimes
            'Return
            Exit Sub
        End If
        If gDateValue(slAirDate) < llSDate Then
            mRestoreSpotTimes
            'Return
            Exit Sub
        End If
    End If
    'Get Copy for spot
    slZone = Trim$(tmEgf.sZone)
    If slZone = "" Then
        slZone = "Oth"
    End If
    
    '8/29/14: Check for Airing vehicle copy (ttp 7064)
    gGetAirCopy tmVef.sType, tmVef.iCode, ilVpfIndex, tmSdf, hmCrf, hmRsf, hmCvf, slZone
    
    ilCopy = mObtainCopy(slZone)
    If Not ilCopy Then
        If llCopyMissingSdfCode <> tmSdf.lCode Then
            'Print #hmMsg, slMsg
            gAutomationAlertAndLogHandler slMsg
            'lbcMsg.AddItem slMsg
            mAddMsgToListBox slMsg
            llCopyMissingSdfCode = tmSdf.lCode
        End If
    End If
    'Later if copy missing- output error and stop creation, delete file


    'Air date
    gObtainYearMonthDayStr slAirDate, True, slFYear, slFMonth, slFDay
    slRecord = slRecord & right$(slFYear, 2) & slFMonth & slFDay & " "
    ''Zone
    'If tmEgf.sZone = "EST" Then
    '    slRecord = slRecord & "03 "
    'ElseIf tmEgf.sZone = "CST" Then
    '    slRecord = slRecord & "04 "
    'ElseIf tmEgf.sZone = "PST" Then
    '    slRecord = slRecord & "06 "
    'Else
    '    slRecord = slRecord & "   "
    'End If
    'Schedule
    slStr = Trim$(tmEgf.sSchedule)
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slRecord = slRecord & slStr & " "
    'Feed Time
    llHour = llSpotTime \ 3600
    llMin = llSpotTime Mod 3600
    llSec = llMin Mod 60
    llMin = llMin \ 60
    slStr = Trim$(str$(llHour))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slRecord = slRecord & slStr
    slStr = Trim$(str$(llMin))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slRecord = slRecord & slStr
    slStr = Trim$(str$(llSec))
    Do While Len(slStr) < 2
        slStr = "0" & slStr
    Loop
    slRecord = slRecord & slStr & " "
    'Bus
    slRecord = slRecord & tmEgf.sBus & " "  'All 5 characters required
    'Cart ID
    If ilCopy Then
        '12/26/13: Suppress Media code
        '4/16/14: If Live, retain Media Code Name
        'If (smEngrHideMediaCode = "Y") Then
        If (smEngrHideMediaCode = "Y") And (Trim$(tmMcf.sName) <> "L") Then
            slRecord = slRecord & tmCif.sName & " " 'retain 5 characters
        Else
            If Trim$(tmCif.sCut) = "" Then
                slRecord = slRecord & Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
            Else
                slRecord = slRecord & Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
            End If
        End If
    Else
        slRecord = slRecord & "C    " & " "
        imCopyMissing = True
    End If
    'Commercial Title
    'If ilCopy And (Trim$(tmCpf.sName) <> "") Then
    '    slShortTitle = UCase$(Left$(tmCpf.sName, 15))
    'Else
    '    slShortTitle = UCase$(Left$(tmChf.sProduct, 15))
    'End If
    If tmChf.iAdfCode <> tmAdf.iCode Then
        tmAdfSrchKey.iCode = tmChf.iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    End If
    slShortTitle = UCase$(Left$(gGetShortTitle(hmVsf, hmClf, hmSif, tmChf, tmAdf, tmSdf), 15))
    Do While Len(slShortTitle) < 15
        slShortTitle = slShortTitle & " "
    Loop
    slRecord = slRecord & slShortTitle & " "
    'Spot length
    If tmSdf.iLen < 60 Then
        slStr = Trim$(str$(tmSdf.iLen))
        Do While Len(slStr) < 4
            slStr = "0" & slStr
        Loop
        slRecord = slRecord & slStr & " "
    ElseIf tmSdf.iLen = 60 Then
        slRecord = slRecord & "0100" & " "
    Else
        llMin = tmSdf.iLen
        llSec = llMin Mod 60
        llMin = llMin \ 60
        slStr = Trim$(str$(llMin))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        slRecord = slRecord & slStr
        slStr = Trim$(str$(llSec))
        Do While Len(slStr) < 2
            slStr = "0" & slStr
        Loop
        slRecord = slRecord & slStr & " "
    End If
    'Vehicle code
    slRecord = slRecord & slCodeStn
    'Set to 52 so that EAS is correct
    Do While Len(slRecord) < 52
        slRecord = slRecord & " "
    Loop
    'ISCI
    If (smEASOrASPOrESPN = "ASP") Or (smEASOrASPOrESPN = "ESPN") Then
        slISCI = Trim$(tmCpf.sISCI)
        slISCI = gFileNameFilter(slISCI)
        Do While Len(slISCI) < 20
            slISCI = slISCI & " "
        Loop
        slRecord = slRecord & slISCI
    End If
    ilRet = 0
    smNewLines(UBound(smNewLines)) = slRecord
    tgSpotSum(UBound(tgSpotSum)).lDate = gDateValue(slAirDate)
    tgSpotSum(UBound(tgSpotSum)).iVefCode = ilVefCode
    tgSpotSum(UBound(tgSpotSum)).iNewIndex = UBound(smNewLines)
    tgSpotSum(UBound(tgSpotSum)).iMnfComp(0) = tmSpot.iMnfComp(0)
    tgSpotSum(UBound(tgSpotSum)).iMnfComp(1) = tmSpot.iMnfComp(1)
    tgSpotSum(UBound(tgSpotSum)).iAdfCode = tmSdf.iAdfCode
    tgSpotSum(UBound(tgSpotSum)).lTime = llSpotTime
    tgSpotSum(UBound(tgSpotSum)).iLen = tmSdf.iLen
    tgSpotSum(UBound(tgSpotSum)).sShortTitle = slShortTitle
    ReDim Preserve tgSpotSum(0 To UBound(tgSpotSum) + 1) As SPOTSUM
    ReDim Preserve smNewLines(0 To UBound(smNewLines) + 1) As String * 72
End Sub
